import urllib.request
from json import loads
import html
import re
import xlsxwriter


def clean(vtitle):
    if len(vtitle) > 0:
        vtitle = html.unescape(vtitle.strip())
        vtitle = re.compile(r'<[^>]+>').sub('', vtitle)
        return ' '.join(str(vtitle).replace('\n', ' ').replace('\r', '').split())
    return ''


class author:
    def __init__(self, name, lastname, first):
        self.first = False
        self.name = clean(str(name).title().replace(" ", ""))
        self.lastname = clean(str(lastname).title())
        self.first = first

    def format_string(self):
        aut = ''
        if len(self.lastname) > 0 and len(self.name) > 0:
            nombres = self.name.split(' ')
            if len(nombres) > 0:
                aut = self.lastname + ' ' + ''.join([x[0] for x in nombres])
            else:
                aut = self.lastname
        return aut


class publicacion:
    def __init__(self,vsource):
        self.source = vsource
        self.primary_source= None
        self.tipo_pub = 'None'
        self.authors = []
        if (vsource[0]['type'] is not None):
            self.tipo_pub=vsource[0]['type']
        self.authors = []
        self.doi = 'No Doi'
        self.ref_url = 'No ref'
        self.title ='No title'
        self.title = clean(vsource[0]['title']['title']['value'])
        self.vol = ''
        self.issn = ''
        self.journal ='No journal'
        if (vsource[0]['journal-title'] is not None):
            self.journal = clean(vsource[0]['journal-title']['value'])
        self.anno = ''
        self.impact = ''
        self.found = False
        self.tipo = {}
        self.EID =''
        self.hasrefdoi=False
        self.hasrefeid=False
        self.add_doi()

    def format_string_doi(self):
        delstring = []
        delstring.append(self.get_autorlist(True))
        delstring.append(self.get_autorlist(False))
        delstring.append(self.anno)
        delstring.append(self.title)
        delstring.append(self.journal)
        delstring.append(self.issn)
        delstring.append(self.doi)
        delstring.append(self.tipo_pub)
        return '|'.join(delstring)
    def format_string_eid(self):
        delstring = []
        delstring.append(self.anno)
        delstring.append(self.title)
        delstring.append(self.journal)
        delstring.append(self.issn)
        delstring.append(self.ref_url)
        delstring.append(self.tipo_pub)
        return '|'.join(delstring)
    def add_volumen(self):
        vol = ""
        if 'volume' in self.primary_source.keys():
            vol = self.primary_source["volume"]
        if 'page' in self.primary_source.keys():
            vol = vol + ':' + self.primary_source["page"]
        elif 'issue' in self.primary_source.keys():
            vol = vol + '(' + self.primary_source["issue"] + ')'
        self.vol = vol

    def add_authors(self):

        for nn in self.primary_source['author']:
            firstauthor = False
            try:
                if (nn.get('given') is not None) or (nn.get('family') is not None):
                    if (nn.get('sequence') is not None):
                        if (nn['sequence'] == 'first'):
                            firstauthor = True
                    if (nn.get('given') is not None) and (nn.get('family') is not None):
                        autor = author(nn['given'], nn['family'], firstauthor)
                    elif nn.get('family') is not None:
                        autor = author('', nn['family'], firstauthor)
                    elif nn.get('given') is not None:
                        autor = author(nn['given'], '', firstauthor)
                    else:
                        autor = 'Unresolved'
            except Exception as e:
                print(str(e) + " - Error en author!")
            self.authors.append(autor)

    def get_autorlist(self, all):
        autlista = []
        for a in self.authors:
            aut = a.format_string()
            if all:
                autlista.append(aut)
            elif a.first:
                autlista.append(aut)
            autlistaclean = []
            [autlistaclean.append(x) for x in autlista if x not in autlistaclean]
        return ', '.join(autlistaclean)

    def get_autorcolab(self):
        autlista = []
        for a in self.authors:
            aut = a.format_string()
            if not a.first:
                autlista.append(aut)
        if len(autlista) > 0:
            return ', '.join(autlista)
        else:
            return self.get_autorlist(True)

    def add_ISSN(self):
        try:
            self.issn = ', '.join(self.primary_source["ISSN"])
        except:
            self.issn = 'No ISSN'

    def add_anno(self):
        try:
            self.anno = str(self.primary_source['published']['date-parts'][0][0])
        except:
            self.anno = '0000'

    def add_EID(self, veid):
        self.EID = clean(veid)
        self.hasrefeid =True

    def set_primary_source(self, vprimary):
        self.primary_source=vprimary
        self.add_authors()
        self.add_volumen()
        self.add_ISSN()
        self.add_anno()
        self.found = True
    def add_doi(self):
        for summary in self.source:
            if 'external-ids' in summary.keys():
                for ids in summary['external-ids']:
                    for id in summary['external-ids']['external-id']:
                        if id.get('external-id-type') is not None:
                            if id.get('external-id-type') == 'doi':
                                url = id.get('external-id-value')
                                if "doi.org" in url:
                                    url = url.split(".org/")[1]
                                urldoi = 'http://dx.doi.org/' + url
                                self.doi = clean(urldoi)
                                self.hasrefdoi=True
                            if id.get('external-id-type') == 'eid':
                                url = summary['url']
                                self.ref_url=clean(url['value'])
                                self.add_EID(id.get('external-id-value'))

def get_doi_orcid(orcidurl):

    req = urllib.request.Request(orcidurl)
    req.add_header('Accept', 'application/json')
    listadoi = []
    path=''
    try:
        with urllib.request.urlopen(req, timeout=15) as f:
            jsonbruto = f.read()
            json = loads(jsonbruto.decode("utf-8"))
            #print(json)
            path=json['orcid-identifier']['path']
            if 'activities-summary' in json.keys():
                activities = json['activities-summary']
                for work in activities['works']['group']:
                    listadoi.append(publicacion(work['work-summary']))
    except ValueError as e:
        print(orcidurl+" "+ str(e)+" An exception occurred")
    except Exception as e:
        print(orcidurl+" "+ str(e)+" An exception occurred")
    return listadoi , path


def load_pubobj(vpub, vjson):
    json_int = loads(vjson.decode("utf-8"))
    try:
        vpub.set_primary_source(json_int)

    except Exception as e:
        print(str(e) + vpub.ref_url + " - Problemas!")
    return vpub

def write_output(vworksheet,vpub, vrow):

    vworksheet.write(vrow, 0,pub_loaded.get_autorlist(True))
    vworksheet.write(vrow, 1,pub_loaded.get_autorlist(False))
    vworksheet.write(vrow, 2,pub_loaded.anno)
    vworksheet.write(vrow, 3,pub_loaded.title)
    vworksheet.write(vrow, 4,pub_loaded.journal)
    vworksheet.write(vrow, 5,pub_loaded.issn)
    vworksheet.write(vrow, 6,pub_loaded.doi)
    vworksheet.write(vrow, 7,pub_loaded.doi)
    vworksheet.write(vrow, 8,pub_loaded.tipo_pub)


vurl = input("Ingresar url ORCID : ")
lista_pub,idarchivo = get_doi_orcid(vurl)
workbook = xlsxwriter.Workbook(idarchivo+'.xlsx')
worksheet = workbook.add_worksheet()
#worksheet.write('A1', 'Hello world')
#archivo = open(idarchivo+".txt", "a")
for ipublicacion in lista_pub:
    if ipublicacion.hasrefdoi:
        req = urllib.request.Request(ipublicacion.doi)
        req.add_header('Accept', 'application/json')
        try:
            with urllib.request.urlopen(req, timeout=15) as f:
                pub_loaded =load_pubobj(ipublicacion, f.read())

                #archivo.write( load_pubobj(ipublicacion, f.read()).format_string_doi() + '\n')
                print(ipublicacion.doi)
        except Exception as e:
            #archivo.write(str(e)+ipublicacion.doi + ' No disponible online' + '\n')
            print(ipublicacion.doi + ' No disponible online')
    elif ipublicacion.hasrefeid:
        #archivo.write('|||' + ipublicacion.format_string_eid() + '\n')
        print(ipublicacion.ref_url)
    else:
       #archivo.write('|||' + ipublicacion.format_string_eid() + '\n')
        0==0
#archivo.close()
workbook.close()
print(len(lista_pub))