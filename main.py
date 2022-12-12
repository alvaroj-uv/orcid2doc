from docx import Document
import pandas
import sqlite3 as sqlite
import urllib.request
from json import loads

def get_doi_orcid(orcidurl):
    req = urllib.request.Request(orcidurl)
    req.add_header('Accept', 'application/json')
    with urllib.request.urlopen(req, timeout=15) as f:
        jsonbruto = f.read()
        json = loads(jsonbruto.decode("utf-8"))
        # print(json)
        listadoi = []
        if 'activities-summary' in json.keys():
            activities = json['activities-summary']
            for works in activities['works']['group']:
                for summary in works['work-summary']:
                    if 'external-ids' in summary.keys():
                        for ids in summary['external-ids']:
                            for id in summary['external-ids']['external-id']:
                                if (id.get('external-id-type')) != None:
                                    if id.get('external-id-type') == 'doi':
                                        url=id.get('external-id-value')
                                        if "doi.org" in url:
                                            url = url.split(".org/")[1]
                                        urldoi = 'http://dx.doi.org/' + url
                                        if urldoi not in listadoi:
                                            listadoi.append(urldoi)
    return listadoi

def update_orcid(filename):
    academicos = pandas.read_excel(filename, sheet_name='Base_Acad')
    academicos.replace(u'\xa0',u'', regex=True, inplace=True)
    academicos['email']=academicos['email'].str.lower()
    academicos.dropna(how='all')
    filtered_academicos = academicos[academicos['Orcid'].notnull()]
    df2 = filtered_academicos[['email', 'Orcid']].copy()
    listapub = []
    for index,row in df2.iterrows():
        print('Actualizando doi publicaciones via Orcid de: '+row["email"])
        listadoi = get_doi_orcid(row["Orcid"])
        print(str(len(listadoi))+' publicaciones')
        for doi in listadoi:
            listapub.append([row["email"],doi])

    publicaciones=pandas.DataFrame(listapub,columns=['email','doi'])

    with pandas.ExcelWriter(filename, mode='a', if_sheet_exists='replace') as writer:
        publicaciones.to_excel(writer, sheet_name='publicaciones',index=False)


def excel_to_bd(filename,vconn):
    update_orcid(filename)
    print('Generando base de datos desde excel')
    academicos = pandas.read_excel(filename, sheet_name='Base_Acad')
    academicos.replace(u'\xa0',u'', regex=True, inplace=True)
    academicos['email']=academicos['email'].str.lower()
    academicos.dropna(how='all')
    academicos.to_sql('base_acad',vconn, if_exists='replace', index=False)

    tesis = pandas.read_excel(filename, sheet_name='tesis_postgrado')
    tesis.replace(u'\xa0',u'', regex=True, inplace=True)
    tesis.dropna(how='all')
    tesis['email']=tesis['email'].str.lower()
    tesis.to_sql('tesis',vconn, if_exists='replace', index=False)

    publicaciones = pandas.read_excel(filename, sheet_name='publicaciones')
    publicaciones.replace(u'\xa0',u'', regex=True, inplace=True)
    publicaciones.dropna(how='all')
    publicaciones['email']=publicaciones['email'].str.lower()
    publicaciones.to_sql('publicaciones',vconn, if_exists='replace', index=False)

    proyectos = pandas.read_excel(filename, sheet_name='proyectos')
    proyectos.replace(u'\xa0',u'', regex=True, inplace=True)
    proyectos.dropna(how='all')
    proyectos['email']=proyectos['email'].str.lower()
    proyectos.to_sql('proyectos',vconn, if_exists='replace', index=False)

    consultoria = pandas.read_excel(filename, sheet_name='consultorias')
    consultoria.replace(u'\xa0',u'', regex=True, inplace=True)
    consultoria.dropna(how='all')
    consultoria['email']=consultoria['email'].str.lower()
    consultoria.to_sql('consultoria',vconn, if_exists='replace', index=False)
    print('Base de datos actualizada')

def add_table_tesis(cell,vconn,id_prof,tipo_tesis):
    def setuptable(table_v):
        table_v.style = 'TableGrid'
        table_v.cell(0, 0).text = 'Año'
        table_v.cell(0, 1).text = 'Autor'
        table_v.cell(0, 2).text = 'Titulo de tesis'
        table_v.cell(0, 3).text = 'Nombre del programa'
        table_v.cell(0, 4).text = 'Institución'

    cell.add_paragraph('Como guía tesis')
    table_t1=cell.add_table(rows=1, cols=5)
    setuptable(table_t1)
    cur = vconn.cursor()
    cur.execute(
        "SELECT b.anno, b.autor, b.titulo, b.programa, b.institucion FROM tesis b where b.tipo =? and b.rol='Guia' and b.email=?",[tipo_tesis,id_prof])
    rows = cur.fetchall()
    for row in rows:
        row_cells = table_t1.add_row().cells
        row_cells[0].text = str(row['anno'])
        row_cells[1].text = row['autor']
        row_cells[2].text = row['titulo']
        row_cells[3].text = row['programa']
        row_cells[4].text = row['institucion']

    cell.add_paragraph('Como co-guía tesis')
    table_t2=cell.add_table(rows=1, cols=5)
    setuptable(table_t2)
    cur.execute(
        "SELECT b.anno, b.autor, b.titulo, b.programa, b.institucion FROM tesis b where b.tipo =? and b.rol='Co-Guia' and b.email=?",[tipo_tesis,id_prof])
    rows = cur.fetchall()
    for row in rows:
        row_cells = table_t2.add_row().cells
        row_cells[0].text = str(row['anno'])
        row_cells[1].text = row['autor']
        row_cells[2].text = row['titulo']
        row_cells[3].text = row['programa']
        row_cells[4].text = row['institucion']
    cur.close()
def add_table_proyectos(cell):
    def setuptable(table_v):
        table_v.style = 'TableGrid'
        table_v.cell(0, 0).text = 'Título'
        table_v.cell(0, 1).text = 'Fuente de financiamiento'
        table_v.cell(0, 2).text = 'Año de adjudicación'
        table_v.cell(0, 3).text = 'Periodo de ejecución'
        table_v.cell(0, 4).text = 'Rol'
    table_t=cell.add_table(rows=1, cols=5)
    setuptable(table_t)


def add_table_consultorias(cell):
    def setuptable(table_v):
        table_v.style = 'TableGrid'
        table_v.cell(0, 0).text = 'Título'
        table_v.cell(0, 1).text = 'Institucion'
        table_v.cell(0, 2).text = 'Año de adjudicación'
        table_v.cell(0, 3).text = 'Periodo de ejecución'
        table_v.cell(0, 4).text = 'Objetivo'
    table_t = cell.add_table(rows=1, cols=5)
    setuptable(table_t)

def add_table_publicaciones(cell,vconn,id_prof):
    def setuptable(table_v):
        table_v.style = 'TableGrid'
        table_v.cell(0, 0).text = 'N°'
        table_v.cell(0, 1).text = 'Autores'
        table_v.cell(0, 2).text = 'Año'
        table_v.cell(0, 3).text = 'Titulo del articulo'
        table_v.cell(0, 4).text = 'Nombre revista'
        table_v.cell(0, 5).text = 'Estado'
        table_v.cell(0, 6).text = 'ISSN'
        table_v.cell(0, 7).text = 'Factor de impacto'
    def setuptablelibro(table_v):
        table_v.style = 'TableGrid'
        table_v.cell(0, 0).text = 'N°'
        table_v.cell(0, 1).text = 'Autores'
        table_v.cell(0, 2).text = 'Año'
        table_v.cell(0, 3).text = 'Titulo del capitulo/libro'
        table_v.cell(0, 4).text = 'Lugar'
        table_v.cell(0, 5).text = 'Editorial'

    def setuptableotro(table_v):
        table_v.style = 'TableGrid'
        table_v.cell(0, 0).text = 'N°'
        table_v.cell(0, 1).text = 'Autores'
        table_v.cell(0, 2).text = 'Año'
        table_v.cell(0, 3).text = 'Titulo de publicación'
        table_v.cell(0, 4).text = 'Lugar'
        table_v.cell(0, 5).text = 'Editorial'
        table_v.cell(0, 6).text = 'Estado'
        table_v.cell(0, 7).text = 'Observaciones'

    def setuptablepatentes(table_v):
        table_v.style = 'TableGrid'
        table_v.cell(0, 0).text = 'N°'
        table_v.cell(0, 1).text = 'Inventores'
        table_v.cell(0, 2).text = 'Nombre Patente'
        table_v.cell(0, 3).text = 'Fecha solicitud'
        table_v.cell(0, 4).text = 'Fecha publicación'
        table_v.cell(0, 5).text = 'N° registro'
        table_v.cell(0, 6).text = 'Estado'


    cell.add_paragraph('Publicaciones indexadas')
    table_t_pub = cell.add_table(rows=1, cols=8)
    setuptable(table_t_pub)
    cur = vconn.cursor()
    cur.execute("select p.email,m.autores,m.anno,m.titulo,ifnull(m.revista,'Revista')||'('||ifnull(m.ref_revista,'')||') ' ||m.doi as revista,m.isbn,m.factor from publicaciones p, master_doi m where p.doi=m.doi and p.email=?",[id_prof])
    rows = cur.fetchall()
    i=1
    for row in rows:
        row_cells = table_t_pub.add_row().cells
        row_cells[0].text = str(i)
        row_cells[1].text = row['autores']
        row_cells[2].text = str(row['anno'])
        row_cells[3].text = row['titulo']
        row_cells[4].text = row['revista']
        row_cells[5].text = 'Publicada'
        row_cells[6].text = row['isbn']
        row_cells[7].text = row['factor']
        i=i+1
    cur.close()

    cell.add_paragraph('Libros y capítulos de libro ')
    table_t = cell.add_table(rows=1, cols=6)
    setuptablelibro(table_t)
    cell.add_paragraph('Otras publicaciones ')
    table_t = cell.add_table(rows=1, cols=8)
    setuptableotro(table_t)
    cell.add_paragraph('Patentes ')
    table_t = cell.add_table(rows=1, cols=7)
    setuptablepatentes(table_t)

def add_data_db_header(filename,vconn):
    excel_to_bd(filename,vconn)
    print('Base actualizada')
    vconn.row_factory = sqlite.Row
    cur = vconn.cursor()
    cur.execute("SELECT b.email,ifnull(b.linea_invest,'') as linea_invest,b.rut,b.nombre,b.tipo,b.profesion||';'||ifnull(b.inst_profesion,'UV')||';'||ifnull(b.pais_profesion,'Sin Info') as profesion, b.max_grado||';'||ifnull(b.institucion_grado,'')||';'||ifnull(round(b.ano_max_grado,0),'')||';'||ifnull(b.pais_grado,'') as grado FROM base_acad b where b.tipo ='Nucleo'")
    rows = cur.fetchall()
    for row in rows:
        print('Escribiendo archivo '+row['nombre'])
        document = Document()
        table = document.add_table(rows=12, cols=2)
        table.style = 'TableGrid'
        hdr_cells = table.rows[0].cells
        table.cell(0,0).text = 'Nombre del académico'
        table.cell(0,1).text = row['nombre']
        table.cell(1,0).text = 'Carácter del vínculo (claustro/núcleo, colaborador o visitante'
        table.cell(1,1).text = row['tipo']
        table.cell(2,0).text = 'Título profesional,  institución, país'
        table.cell(2,1).text = row['profesion']
        table.cell(3,0).text = 'Grado académico máximo (especificar área disciplinar), institución, año de graduación y país '
        table.cell(3,1).text = row['grado']
        table.cell(4,0).text = 'Línea(s) de investigación'
        table.cell(4,1).text = row['linea_invest']

        table.cell(5,0).text = 'Tesis de magíster  dirigidas en los últimos 10 años (finalizadas)'
        add_table_tesis(table.cell(5,1),vconn,row['email'],'Magister')

        table.cell(6,0).text = 'Tesis de doctorado dirigidas en los últimos 10 años (finalizadas)'
        add_table_tesis(table.cell(6, 1),vconn,row['email'],'Doctorado')

        table.cell(7,0).text = 'Listado de publicaciones'
        add_table_publicaciones(table.cell(7,1),vconn,row['email'])

        table.cell(8,0).text = 'Listado de proyectos de investigación  en los últimos 10 años'
        add_table_proyectos(table.cell(8,1))

        table.cell(9,0).text = 'Listado de proyectos de intervención, innovación y/o desarrollo tecnológico'
        add_table_proyectos(table.cell(9,1))

        table.cell(10,0).text = 'Listado de consultorías y/o  asistencias técnicas, en calidad de responsable, en los últimos 10 años'
        add_table_consultorias(table.cell(10,1))

        document.save('./output/'+row['rut']+'_'+row['nombre']+'.docx')

conn = sqlite.connect('./bd_academic.sqlite')
add_data_db_header('Base_Academicos.xlsx',conn)