import html
import re


def clean(vtitle):
    if len(vtitle) > 0:
        vtitle = html.unescape(vtitle.strip())
        vtitle = re.compile(r'<[^>]+>').sub('', vtitle)
        return ' '.join(str(vtitle).replace('\n', ' ').replace('\r', '').split())
    return ''

class author:
    def __init__(self, name, lastname,first):
        self.first = False
        self.name = clean(str(name).title().replace(" ",""))
        self.lastname = clean(str(lastname).title())
        self.first = first


class publicacion:


    def __init__(self, title, doi,journal):
        tipopub = set(("article", "book"))
        self.authors = []
        self.doi = clean(doi)
        self.title = clean(title)
        self.vol = ''
        self.issn = ''
        self.journal = clean(journal)
        self.anno = ''
        self.impact = ''
        self.found = False
        self.tipo = {}


    def add_volumen(self,vjson):
        vol = ""
        if 'volume' in vjson.keys():
            vol = vjson["volume"]
        if 'page' in vjson.keys():
            vol = vol + ':' + vjson["page"]
        elif 'issue' in vjson.keys():
            vol = vol + '(' + vjson["issue"] + ')'
        self.vol = vol
    def add_authors(self, authorlist):
        for nn in authorlist:
            firstauthor = False
            try:
                if (nn.get('given') is not None) or (nn.get('family') is not  None):
                    if (nn.get('sequence') is not None):
                        if (nn['sequence'] == 'first'):
                            firstauthor = True
                    if (nn.get('given') is not None) and (nn.get('family') is not  None):
                        autor = author(nn['given'], nn['family'], firstauthor)
                    elif nn.get('family') is not  None:
                        autor = author('', nn['family'], firstauthor)
                    elif nn.get('given') is not  None:
                        autor = author(nn['given'], '', firstauthor)
                    else:
                        autor = 'Unresolved'
            except Exception as e:
                print(str(e) + " - Error en author!")
            self.authors.append(autor)

    def get_autorlist(self, all):
        autlista = []
        for a in self.authors:
            aut = a.lastname
            if len(a.lastname) > 0 and len(a.name) > 0:
                nombres = a.name.split(' ')
                if len(nombres) > 0:
                    aut = a.lastname + ' ' + ''.join([x[0] for x in nombres])
                else:
                    aut = a.lastname
            if all:
                autlista.append(aut)
            elif a.first:
                autlista.append(aut)
            autlistaclean=[]
            [autlistaclean.append(x) for x in autlista if x not in autlistaclean]
        return ', '.join(autlistaclean)

    def get_autorcolab(self):
        autlista = []
        for a in self.authors:
            aut = a.lastname
            if len(a.lastname) > 0 and len(a.name) > 0:
                nombres = a.name.split(' ')
                if len(nombres) > 0:
                    aut = a.lastname + ' ' + ''.join([x[0] for x in nombres])
                else:
                    aut = a.lastname
            if not a.first:
                autlista.append(aut)
        if len(autlista) > 0:
            return ', '.join(autlista)
        else:
            return self.get_autorlist(True)
