"""Local aggregate generator. Never copy input workbooks or record-level outputs to the public repo.

Adapted from Jorge Arango's trayectorias_musica/_build_data.py (permission confirmed by requester).
Python + vanilla JS/SVG; no framework. All joins use source + record ID + employer_key.
"""
from pathlib import Path
from collections import Counter,defaultdict
from statistics import mean,median
from itertools import product
from datetime import date
import argparse,json,re,unicodedata,hashlib,math,csv,colorsys
from openpyxl import load_workbook

K=2
COHORTS=['todas','2017–2019','2020–2021','2022–2023','2024–2025','2026 provisional']
UNKNOWN='No determinado'
PROFESSIONAL_AREAS=['Conservación y ecología','Biotecnología y laboratorio','Bioinformática y análisis de datos','Educación y divulgación científica','Otros']
GROUPED={'employer_name':'Otros empleadores agrupados','title_normalized':'Otros cargos agrupados'}

def sin_tildes(s):
 return ''.join(c for c in unicodedata.normalize('NFD',str(s)) if unicodedata.category(c)!='Mn')
def clean(s):return ' '.join(str(s or '').replace('\u200b','').split())
def token(s):return sin_tildes(clean(s)).casefold()
def is_unknown(s):return token(s) in ['','no determinado','no determinada','sin categoria']
def cohort(d):
 y=int(str(d)[:4]);return COHORTS[1] if y<=2019 else COHORTS[2] if y<=2021 else COHORTS[3] if y<=2023 else COHORTS[4] if y<=2025 else COHORTS[5]
def read(p):return json.loads(Path(p).read_text(encoding='utf-8'))
def write(p,d):
 Path(p).parent.mkdir(parents=True,exist_ok=True);Path(p).write_text(json.dumps(d,ensure_ascii=False,allow_nan=False,separators=(',',':')),encoding='utf-8')
def unpack(t):return [dict(zip(t['headers'],r)) for r in t['rows']]
def tab(rows,headers=None):
 headers=headers or list(dict.fromkeys(k for r in rows for k in r));return {'headers':headers,'rows':[[r.get(h) for h in headers] for r in rows]}
def excel(path,sheet):
 w=load_workbook(path,read_only=True,data_only=True);ws=w[sheet] if isinstance(sheet,str) else w.worksheets[sheet];ws.reset_dimensions();it=iter(ws.values);h=next(it);rs=[dict(zip(h,r)) for r in it if any(v is not None for v in r)];w.close();return rs
TITLE_ALIASES={
 'biologa':'Biólogo','biologo':'Biólogo','biologist':'Biólogo',
 'research assistant':'Asistente de investigación','asistente de investigacion':'Asistente de investigación',
 'young researcher':'Joven investigador','joven investigador':'Joven investigador','joven investigadora':'Joven investigador',
 'researcher':'Investigador','investigadora':'Investigador','investigador':'Investigador',
 'scientific researcher':'Investigador científico','research scientist':'Investigador científico',
 'junior researcher':'Investigador junior','investigador junior':'Investigador junior',
 'associate researcher':'Investigador asociado','investigadora asociada':'Investigador asociado',
 'visiting researcher':'Investigador visitante','visiting research scientist':'Investigador científico visitante',
 'doctoral student':'Estudiante de doctorado','phd student':'Estudiante de doctorado','estudiante de doctorado':'Estudiante de doctorado',
 'masters student':'Estudiante de maestría','research master student':'Estudiante de maestría','estudiante de master cientifico':'Estudiante de maestría',
 'teaching assistant':'Asistente de docencia','graduate research and teaching assistant':'Asistente de investigación y docencia de posgrado',
 'graduate research assistant':'Asistente de investigación de posgrado','student assistant':'Asistente estudiantil',
 'research intern':'Practicante de investigación','intern':'Practicante','internship trainee':'Practicante','student intern':'Practicante','pasante':'Practicante',
 'summer intern':'Practicante de verano','laboratory assistant':'Asistente de laboratorio',
 'data scientist':'Científico de datos','cientifico de datos':'Científico de datos','cientifica de datos':'Científico de datos',
 'data analyst':'Analista de datos','analista de datos':'Analista de datos','data engineer':'Ingeniero de datos','ingeniero de datos':'Ingeniero de datos',
 'bioinformatician':'Bioinformático','bioinformatician junior':'Bioinformático junior','bioinformatics engineer':'Ingeniero bioinformático',
 'wildlife biologist':'Biólogo de vida silvestre','biologo de vida silvestre':'Biólogo de vida silvestre',
 'project manager':'Administrador de proyectos','customer care':'Servicio de atención al cliente',
 'university lecturer':'Profesor universitario','profesor universitario':'Profesor universitario',
 'docencia':'Docente','docente':'Docente','maestro':'Maestro','maestra':'Maestro',
 'biologa investigadora':'Biólogo investigador','biologo investigador':'Biólogo investigador',
 'profesional ambiental':'Profesional ambiental','analista tecnico':'Analista técnico',
 'mastozoóloga':'Mastozoólogo','masto zoologa':'Mastozoólogo','mastozoóloga':'Mastozoólogo',
}
def normalize_title(value):
 text=clean(value);key=token(text)
 if key in TITLE_ALIASES:return TITLE_ALIASES[key]
 if key=='mastozoologa' or key=='mastozoóloga' or key=='masto zoologa':return 'Mastozoólogo'
 # Gender variants only; preserve specialisms, seniority and composite roles.
 for feminine,masculine in [('bióloga','Biólogo'),('investigadora','investigador'),('consultora','consultor'),('coordinadora','coordinador'),('administradora','administrador'),('mentora','mentor'),('ingeniera','ingeniero'),('científica','científico'),('educadora','educador'),('profesora','profesor'),('técnica','técnico'),('directora','director')]:
  text=re.sub(r'\b'+feminine+r'\b',masculine,text,flags=re.I)
 return text

def categorical_colors(labels,previous=None,priority=None,neighbors=None):
 """Append-only registry maximizing distance among labels sharing a chart."""
 def rgb(h,s,l):return tuple(round(255*v) for v in colorsys.hls_to_rgb(h/360,l,s))
 def hx(c):return '#'+''.join(f'{v:02x}' for v in c)
 def lab(c):
  v=[x/255 for x in c];v=[x/12.92 if x<=.04045 else ((x+.055)/1.055)**2.4 for x in v]
  xyz=[sum(a*b for a,b in zip(v,row))/ref for row,ref in [([.4124,.3576,.1805],.95047),([.2126,.7152,.0722],1),([.0193,.1192,.9505],1.08883)]]
  f=[x**(1/3) if x>.008856 else 7.787*x+16/116 for x in xyz];return [116*f[1]-16,500*(f[0]-f[1]),200*(f[1]-f[2])]
 def decode(s):return tuple(int(s[i:i+2],16) for i in [1,3,5])
 palette={k:v for k,v in (previous or {}).items() if k in labels};assigned={k:(lab(decode(v[0])),lab(decode(v[1]))) for k,v in palette.items()};used=list(assigned.values());neighbors=neighbors or {}
 candidates=[]
 for h,s,l in product(range(0,360,6),[.60,.78,.94],[.34,.42,.50,.58]):
  pair=[hx(rgb(h,s,l)),hx(rgb(h,s,min(l+.20,.80)))];labs=(lab(decode(pair[0])),lab(decode(pair[1])))
  if pair not in palette.values():candidates.append([pair,labs,float('inf')])
 def distance(a,b):return min(sum((x-y)**2 for x,y in zip(a[i],b[i])) for i in [0,1])
 for c in candidates:
  if used:c[2]=min(distance(c[1],u) for u in used)
 order=list(dict.fromkeys([label for label in (priority or []) if label in labels]+sorted(labels,key=lambda label:(-len(neighbors.get(label,set())),label))))
 for label in [label for label in order if label not in palette]:
  light_used={v[0] for v in palette.values()};dark_used={v[1] for v in palette.values()}
  candidates=[c for c in candidates if c[0][0] not in light_used and c[0][1] not in dark_used]
  peers=[assigned[k] for k in neighbors.get(label,set()) if k in assigned]
  if peers:chosen=max(candidates,key=lambda c:(min(distance(c[1],peer) for peer in peers),c[2]))
  else:chosen=min(candidates,key=lambda c:sum((x-y)**2 for x,y in zip(c[1][0],lab((37,99,235)))))
  palette[label]=chosen[0];assigned[label]=chosen[1];used.append(chosen[1]);candidates.remove(chosen)
  for c in candidates:c[2]=min(c[2],distance(c[1],chosen[1]))
 return palette
def size_of(a,b):
 """A numeric cell labelled published is not sufficient: inspect its evidence and scope."""
 if token(a['Empleador normalizado'])=='profesional independiente':
  return dict(size_estimate=1,size_method='Convención analítica',size_evidence_type='Convención: independiente = 1',size_confidence='Convención',size_low=False,size_scope='Profesional independiente',size_reference_date='Convención analítica',size_evidence='',size_warning='No es una plantilla observada',size_pending=a['Estado'])
 r=b or a
 if b:
  published=b['Cifra publicada (personas)'];rng=b['Rango sugerido'];lo=b['Límite inferior'];hi=b['Límite superior'];kind=b['Tipo de evidencia'];confidence=b['Certeza'];scope=b['Alcance'];ref=b['Fecha de referencia'];url=b['URL fuente'];basis=b['Fundamento / advertencia'];status=b['Estado']
 else:
  published=a['N.º empleados (cifra publicada)'];rng=a['RANGO PROPUESTO'];lo=a['Límite inferior'];hi=a['Límite superior'];confidence=a['Certeza tamaño'];scope=a['Alcance de tamaño'];ref=a['Fecha de referencia'];url=a['URL evidencia tamaño'];basis=a['Fundamento tamaño'];status=a['Estado'];kind='Rango publicado' if rng else 'Sin estimación'
 # Some supplied rows publish a band but leave its numeric bound cells blank.
 if rng and (lo is None or hi is None):
  numbers=re.findall(r'\d[\d.,]*',str(rng))
  numbers=[int(re.sub(r'[.,]','',v)) for v in numbers]
  if len(numbers)==2:
   if lo is None:lo=numbers[0]
   if hi is None:hi=numbers[1]
  elif len(numbers)==1 and 'mas de' in token(rng) and lo is None:lo=numbers[0]+1
 text=token(basis);scope_text=token(scope)
 uncertain=any(w in text+' '+scope_text for w in ['verificar','confirmar','no comparable','probable','si corresponde','si el ','si la ','no comprobado','no determinada'])
 # Comparable headcounts explicitly supported by the supplied evidence only.
 exact=(kind=='Cifra publicada' and published is not None and not uncertain)
 if not b and published is not None:
  exact=(confidence=='Alta' or ('headcount' in text and not uncertain)) and not any(w in text for w in ['aproxima','rango','verificar'])
 value=None;method='Sin estimación'
 if exact:
  value=float(published);method='Cifra publicada respaldada';kind='Cifra publicada' if not 'secundaria' in text else 'Cifra publicada, referencia secundaria'
 elif lo is not None and hi is not None:
  value=(float(lo)+float(hi))/2;method='Punto medio del rango'
 elif lo is not None and (float(lo)>=10000 or 'mas de 10' in token(rng)):
  value=10001;method='Convención de rango abierto >10.000'
 if value is not None and not b and not exact and any(w in text for w in ['probable','aproxima','verificar','headcount','academicos +']):kind='Referencia indirecta'
 return dict(size_estimate=value,size_method=method,size_evidence_type=kind,size_confidence=confidence or 'Indeterminada',size_low=(confidence=='Baja' or uncertain or kind in ['Hipótesis sin verificar','Referencia indirecta']),size_scope=scope,size_reference_date=str(ref) if ref else None,size_evidence=url,size_warning=basis,size_pending=status,size_published_original=published,size_lower=lo,size_upper=hi,size_range=rng)

class Disclosure:
 """Two-person cells, exact values, grouped remainders and filtered-query guards.

 Reference source/PG views are explicitly released aggregate tables. Narrower
 cohort/combined-filter versions are blocked when differencing the same metric
 would isolate one contributor. A blocked partition is withheld as a whole,
 so its complement cannot reconstruct a suppressed category from the total.
 """
 def __init__(self):self.cells=[];self.index={};self.stats=Counter();self.context='';self.reference=False;self.partitions=[];self.compositions=[]
 def add(self,rows,scalar=None,minimum=K,family='total',reference=None):
  sig=tuple(sorted((r['uid'],r['gid']) for r in rows));key=(family,self.context,sig,scalar,minimum)
  if key in self.index:return self.index[key]
  i=len(self.cells);self.index[key]=i
  people=frozenset(g for _,g in sig if g)
  self.cells.append({'sig':sig,'people':people,'n':len(sig),'scalar':scalar,'min':minimum,'blocked':bool(sig) and len(people)<minimum,'family':family,'context':self.context,'reference':self.reference if reference is None else reference})
  return i
 def protect(self):
  families=defaultdict(list)
  for c in self.cells:
   c['records']=frozenset(u for u,g in c['sig']);families[c['family']].append(c)
  for family,cells in families.items():
   accepted=[]
   for c in sorted(cells,key=lambda c:(not c['reference'],len(c['context']))):
    if c['blocked']:continue
    if not c['reference']:
     for prior in accepted:
      difference=c['records']^prior['records']
      if difference:
       owners=dict(c['sig']+prior['sig']);people={owners[u] for u in difference if owners[u]}
       if len(people)<K:c['blocked']=True;break
    if not c['blocked']:accepted.append(c)
  for indices in self.partitions:
   if any(self.cells[i]['blocked'] for i in indices):
    for i in indices:self.cells[i]['blocked']=True
  for components,statistic in self.compositions:
   if any(self.cells[i]['blocked'] for i in components):
    for i in components:self.cells[i]['blocked']=True
    if statistic is not None:self.cells[statistic]['blocked']=True
  for c in self.cells:self.stats['suppressed' if c['blocked'] else 'released']+=1
 def release(self,i):
  c=self.cells[i]
  if c['blocked']:return None
  if c['scalar'] is not None:return c['scalar']
  return {'exp':c['n'],'grad':len(c['people'])}

def partition(rows,field,disclosure,family,show_unknown=True,ensure=None):
 """Exhaustive disjoint record partition; unique graduate memberships may overlap."""
 groups=defaultdict(list);small=[];grouped_label=GROUPED.get(field,'Otros grupos agrupados');small_labels=set()
 for r in rows:
  label=UNKNOWN if is_unknown(r.get(field)) else str(r[field])
  if field=='employer_name' and (r.get('employer_personal') or token(label)=='profesional independiente'):small.append(r)
  else:groups[label].append(r)
 for label in list(groups):
  if len({r['gid'] for r in groups[label] if r['gid']})<K and groups[label]:small.extend(groups.pop(label));small_labels.add(label)
 for label in ensure or []:
  if label not in small_labels:groups.setdefault(label,[])
 if small:groups[grouped_label]=small
 indices=[];result=[]
 for label,rs in sorted(groups.items()):
  i=disclosure.add(rs,family=family+':'+label);indices.append(i)
  if show_unknown or label!=UNKNOWN:result.append({'label':label,'value':{'$':i},'grouped':label==grouped_label})
 disclosure.partitions.append(indices)
 return result

def main():
 ap=argparse.ArgumentParser(description=__doc__)
 for arg in ['review','employers','sizes','practice','comfama','private-output']:ap.add_argument('--'+arg,type=Path,required=True)
 ap.add_argument('--out',type=Path,default=Path(__file__).parent/'datos/biologia_datos.json')
 args=ap.parse_args();out=args.private_output
 if out.resolve().is_relative_to(Path(__file__).resolve().parent):ap.error('--private-output must be outside the public repository')
 out.mkdir(parents=True,exist_ok=True)
 review=read(args.review);jobs=unpack(next(s for s in review['sheets'] if s['name']=='Final'));grads=unpack(review['sheets'][0]);gmap={r['graduate_id']:r for r in grads};candidates=unpack(review['sheets'][1]);com=read(args.comfama);episodes=unpack(com['tables']['Episodes']);practice=read(args.practice)
 employers=excel(args.employers,'Empleadores');sizes=excel(args.sizes,1);links=excel(args.employers,'Correspondencias')
 emap={r['employer_key']:r for r in employers};smap={r['employer_key']:r for r in sizes};mapping={(r['Fuente'],str(r['ID de registro'])):r for r in links}
 assert len(emap)==len(employers) and len(smap)==len(sizes) and len(mapping)==len(links)
 assert set(smap)<=set(emap)
 personal_keys={mapping['Comfama',r['episode_id']]['employer_key'] for r in episodes if token(r['employer_document_type'])!='nit'}
 registry={}
 for k,a in emap.items():registry[k]={'employer_key':k,'employer_name':a['Empleador normalizado'],'economic_sector':a['SECTOR PROPUESTO'] or UNKNOWN,'economic_sector_confidence':a['Certeza sector'],'economic_sector_evidence':a['Fundamento sector'],'employer_status':a['Estado'],'employer_personal':k in personal_keys,**size_of(a,smap.get(k)),'audit_original':a,'size_audit_original':smap.get(k)}
 rows=[]
 for src,items in [('LinkedIn',jobs),('Comfama',episodes)]:
  for r in items:
   eid=r['employment_id'] if src=='LinkedIn' else r['episode_id'];m=mapping[src,eid];a=registry[m['employer_key']]
   area=r['sector'] if src=='LinkedIn' else r['category'];assert area==m['Área profesional (original)'],('Professional classification drift',eid)
   rows.append(dict(r,**{k:v for k,v in a.items() if k not in ['audit_original','size_audit_original']},uid=src+':'+eid,gid=r['graduate_id'],source=src,cohort=cohort(gmap[r['graduate_id']]['graduation_date']),professional_area=area or UNKNOWN,title_normalized=normalize_title(r['title']) if src=='LinkedIn' else UNKNOWN,country=r.get('country_normalized') or UNKNOWN,city=r.get('city_normalized') or UNKNOWN,postgraduate=r.get('exclude_postgraduate_scenario',False),position=r.get('position_order'),area_basis='Cargo documentado' if src=='LinkedIn' else 'Proxy de actividad del empleador; cargo no observado'))
 li=[r for r in rows if r['source']=='LinkedIn'];counts=Counter(r['gid'] for r in li);fixed={g for g,n in counts.items() if n in [5,6]}
 for r in rows:r['fixed_sample']=r['gid'] in fixed
 assert len(jobs)==323 and len(episodes)==57 and len(rows)==380
 # Private record outputs retain full evidence and uncertainty; never served.
 write(out/'employer_registry.json',list(registry.values()));write(out/'records_enriched.json',rows)
 write(out/'employer_join_report.json',{'employers':len(registry),'complementary_sizes':len(smap),'mapped_current':len(rows),'unselected_mapping_records':[m for key,m in mapping.items() if key not in {(r['source'],r['uid'].split(':',1)[1]) for r in rows}],'size_methods':dict(Counter(a['size_method'] for a in registry.values())),'size_evidence_types':dict(Counter(a['size_evidence_type'] for a in registry.values())),'missing_size_employers':[a['employer_key'] for a in registry.values() if a['size_estimate'] is None]})
 # Enrich current selected rows and all mapped historical Experiences without removing originals.
 extra=['employer_key','economic_sector','economic_sector_confidence','economic_sector_evidence','employer_status','professional_area','title_normalized','size_estimate','size_method','size_evidence_type','size_confidence','size_low','size_scope','size_reference_date','size_evidence','size_warning','size_pending']
 for sheet in review['sheets']:
  if sheet['name'] not in ['Final','Experiences']:continue
  rs=unpack(sheet)
  for r in rs:
   m=mapping.get(('LinkedIn',r['employment_id']));a=registry[m['employer_key']] if m else {}
   r.update({k:a.get(k) for k in extra});r['professional_area']=r['sector'];r['title_normalized']=normalize_title(r['title'])
  sheet.update(tab(rs,sheet['headers']+[h for h in extra if h not in sheet['headers']]))
 review['sheets'].append(dict(name='Employer audit',**tab([{k:(json.dumps(v,ensure_ascii=False) if isinstance(v,dict) else v) for k,v in a.items()} for a in registry.values()])))
 write(out/'review_enriched.json',review)
 with (out/'employer_sizes_pending.csv').open('w',encoding='utf-8-sig',newline='') as f:
  fields=['employer_key','employer_name','size_estimate','size_method','size_evidence_type','size_confidence','size_low','size_scope','size_reference_date','size_evidence','size_warning','size_pending'];w=csv.DictWriter(f,fieldnames=fields,extrasaction='ignore');w.writeheader();w.writerows(registry.values())
 # Display-only equivalence mapping. Original title and area remain unchanged.
 title_labels={}
 for r in li:title_labels.setdefault(token(r['title_normalized']),r['title_normalized'])
 for r in li:r['title_normalized']=title_labels[token(r['title_normalized'])]
 disclosure=Disclosure();views={};exact={};private_long=[]
 def cell(rs,family='total',reference=None):return {'$':disclosure.add(rs,family=family,reference=reference)}
 def scalar(rs,v,family):return {'$':disclosure.add(rs,v,family=family)} if rs else None
 def distribution(rs,field,unknown=True,family=None):return partition(rs,field,disclosure,family or field,unknown,PROFESSIONAL_AREAS if field=='professional_area' else None)
 def subset(co,comf,pg,other):return [r for r in rows if (co=='todas' or r['cohort']==co) and (comf or r['source']=='LinkedIn') and not(pg and r['postgraduate']) and not(other and r['professional_area']=='Otros')]
 for co,comf,pg,other,nd in product(COHORTS,[0,1],[0,1],[0,1],[0,1]):
  key='|'.join(map(str,[co,comf,pg,other,nd]));rs=subset(co,comf,pg,other);linkedin=[r for r in rs if r['source']=='LinkedIn']
  disclosure.context=key;disclosure.reference=co=='todas' and not other
  v={'total':cell(rs),'linkedin':cell(linkedin,'linkedin_total'),'comfama':cell([r for r in rs if r['source']=='Comfama'],'comfama_total'),'economic_sector':distribution(rs,'economic_sector',nd),'professional_area':distribution(rs,'professional_area',nd),'employers':distribution(rs,'employer_name',nd),'titles':distribution(linkedin,'title_normalized',nd),'countries':distribution(rs,'country',nd),'cities':distribution(rs,'city',nd),'details':{},'longitudinal':{}}
  for dim in ['economic_sector','professional_area']:
   v['details'][dim]={}
   for item in v[dim]:
    if item.get('grouped'):continue
    group=[r for r in rs if (UNKNOWN if is_unknown(r[dim]) else r[dim])==item['label']];title_rows=[r for r in group if r['source']=='LinkedIn'];v['details'][dim][item['label']]={'employers':distribution(group,'employer_name',nd,f'{dim}:{item["label"]}:employers'),'titles':distribution(title_rows,'title_normalized',nd,f'{dim}:{item["label"]}:titles'),'linkedin_total':cell(title_rows,f'{dim}:{item["label"]}:titles_total')}
  # First and latest refer to immutable original selected positions.
  for which,positions in [('first',{g:1 for g in counts}),('last',dict(counts))]:
   selected=[r for r in linkedin if r['position']==positions[r['gid']]]
   v[which]={'total':cell(selected,which+'_total'),'sectors':distribution(selected,'economic_sector',nd,which+'_economic_sector'),'areas':distribution(selected,'professional_area',nd,which+'_professional_area')}
  # Counts of original experiences are computed before employment filters.
  originals=[r for r in li if co=='todas' or r['cohort']==co]
  represented={r['gid'] for r in originals}
  count_rows=[dict(r,original_experience_count=str(counts[r['gid']])) for r in originals if r['position']==1]
  v['experience_counts']=distribution(count_rows,'original_experience_count',nd)
  duration=[]
  for r in linkedin:
   lo=r.get('effective_start_min');hi=r.get('effective_end_max')
   if lo and hi:
    months=(date.fromisoformat(hi[:10])-date.fromisoformat(lo[:10])).days/30.4375
    bucket='<6 meses' if months<6 else '6–11 meses' if months<12 else '12–23 meses' if months<24 else '24–47 meses' if months<48 else '48 meses o más'
   else:bucket=UNKNOWN
   duration.append(dict(r,duration_bucket=bucket))
  v['duration']=distribution(duration,'duration_bucket',nd)
  known=[r for r in rs if r['country']!=UNKNOWN];foreign=[r for r in known if r['country']!='Colombia']
  v['international']={'known':cell(known,'international:known'),'foreign':cell(foreign,'international:foreign'),'unknown':cell([r for r in rs if r['country']==UNKNOWN],'international:unknown')}
  for balanced,low in product([0,1],[0,1]):
   disclosure.context=key+f'|{balanced}|{low}'
   disclosure.reference=co=='todas' and not pg and not other and not low
   seq=[]
   for pos in range(1,7):
    eligible=[r for r in originals if r['position']==pos and (not balanced or r['gid'] in fixed)]
    selected=[r for r in linkedin if r['position']==pos and (not balanced or r['gid'] in fixed)]
    sized=[r for r in selected if r['size_estimate'] is not None and not(low and r['size_low'])]
    means=mean(r['size_estimate'] for r in sized) if sized else None;med=median(r['size_estimate'] for r in sized) if sized else None
    missing=[r for r in selected if r['size_estimate'] is None];low_excluded=[r for r in selected if r['size_estimate'] is not None and low and r['size_low']];filtered=[r for r in eligible if r['uid'] not in {s['uid'] for s in selected}]
    prefix=f'position:{pos}'
    row={'position':pos,'eligible':cell(eligible,prefix+':eligible',reference=co=='todas'),'retained':cell(selected,prefix+':retained'),'usable':cell(sized,prefix+':usable'),'missing_size':cell(missing,prefix+':missing'),'excluded_by_filters':cell(filtered,prefix+':excluded'),'excluded_low_size':cell(low_excluded,prefix+':low'),'mean':scalar(sized,means,prefix+':mean'),'mean_status':'available' if sized else 'missing','sectors':distribution(selected,'economic_sector',nd,prefix+':sectors'),'areas':distribution(selected,'professional_area',nd,prefix+':areas')}
    disclosure.compositions.append(([row[f]['$'] for f in ['usable','missing_size','excluded_low_size']],row['mean']['$'] if row['mean'] else None));seq.append(row)
    if not comf and not nd:private_long.append(dict(cohort=co,exclude_postgraduate=pg,exclude_other=other,fixed_sample=balanced,exclude_low_size=low,position=pos,eligible=len(eligible),retained=len(selected),usable=len(sized),coverage=len(sized)/len(eligible) if eligible else None,mean=means,median=med))
   v['longitudinal'][f'{balanced}|{low}']=seq
  views[key]=v
  exact[key]={'linkedin':len(linkedin),'comfama':len(rs)-len(linkedin),'observations':len(rs),'graduates':len({r['gid'] for r in rs}),'sectors':dict(Counter(r['economic_sector'] for r in rs)),'areas':dict(Counter(r['professional_area'] for r in rs))}
 # Practices and M0 obey only compatible filters: cohort; PG for continuity.
 pt=unpack(practice['tables']['Practice Trace']);pairs=unpack(practice['tables']['Continuity']);m0=unpack(practice['tables']['M0 Trace']);formation={}
 for co,pg in product(COHORTS,[0,1]):
  disclosure.context=f'formation|{co}|{pg}';disclosure.reference=co=='todas'
  chosen=[r for r in pt if r['Included in aggregates'] and (co=='todas' or r['Cohort']==co)]
  pr=[dict(uid=r['Practice record id'],gid=r['Graduate id'] or None,entity=r['Entity canonical'],country=r['Country reviewed'] or r['Country group'] or UNKNOWN,city=r['City reviewed'] or r['City (original)'] or UNKNOWN,scope=r['Geographic scope']) for r in chosen]
  cv=[]
  for scope in ['All entities','EAFIT','Other entities']:
   group=[r for r in chosen if scope=='All entities' or (('eafit' in token(r['Entity canonical']))==(scope=='EAFIT'))]
   # Documentary coverage denominators belong to the closed original Final;
   # the PG sensitivity removes employment links, not the baseline population.
   people={r['Graduate id'] for r in group if r['Graduate id']};in_final={r['gid'] for r in li};accepted=[r for r in pairs if r['Accepted all'] and not(pg and r['Exclude postgraduate scenario']) and (co=='todas' or r['Cohort']==co) and (scope=='All entities' or r['Entity scope']==scope)]
   def person_rows(ids,prefix):return [dict(uid='person:'+g,gid=g) for g in ids]
   cv.append({'scope':scope,'practice_graduates':cell(person_rows(people,'practice'),'continuity:'+scope+':practice'),'jointly_observed':cell(person_rows(people&in_final,'joint'),'continuity:'+scope+':joint'),'graduates':cell(person_rows({r['Graduate id'] for r in accepted},'continuity'),'continuity:'+scope+':people'),'matches':cell([dict(uid=r['Employment id'],gid=r['Graduate id']) for r in accepted],'continuity:'+scope+':matches')})
  chosen_m0=[r for r in m0 if r['Selected latest for graduate'] and (co=='todas' or r['Cohort']==co)]
  mr=[dict(uid='m0:'+str(r['Graduate ID']),gid=str(r['Graduate ID']),paid=r['Paid work status'],offer=r['Offer status'],accepted=r['Acceptance status'],working=r['Practice-entity work status'],timing=r['Response timing vs graduation']) for r in chosen_m0]
  # Apply the same all-cohort offer coverage rule already closed in the source:
  # early-cohort form did not establish offer-item comparability.
  for rr in mr:
   if cohort(gmap[rr['gid']]['graduation_date'])==COHORTS[1]:rr['offer']='Missing';rr['accepted']='Missing';rr['working']='Missing'
  formation[f'{co}|{pg}']={'total':cell(pr,'practice_total'),'entities':partition(pr,'entity',disclosure,'practice:entity'),'countries':partition(pr,'country',disclosure,'practice:country'),'cities':partition(pr,'city',disclosure,'practice:city'),'geography':partition(pr,'scope',disclosure,'practice:scope'),'continuity':cv,'m0':{'respondents':cell(mr,'m0:total'),**{field:partition(mr,field,disclosure,'m0:'+field) for field in ['paid','offer','accepted','working','timing']}}}
 # Coverage does not change when a recorded episode is excluded by a chart filter.
 coverage={};approved={r['graduate_id']:r['candidate_id'] for r in candidates if r['identity_status']=='approved'};selected_gids=set(counts)
 experience_by_candidate=defaultdict(list)
 for r in unpack(next(s for s in review['sheets'] if s['name']=='Experiences')):experience_by_candidate[r['graduate_id'],r['candidate_id']].append(r)
 cov=unpack(com['tables']['Coverage'])
 for co in COHORTS:
  disclosure.context='coverage|'+co;disclosure.reference=co=='todas'
  cc=[dict(uid='coverage:'+r['graduate_id'],gid=r['graduate_id'],group=r['coverage_group'],comfama=bool(r['comfama'])) for r in cov if co=='todas' or cohort(gmap[r['graduate_id']]['graduation_date'])==co]
  absence=[]
  for r in cc:
   if r['gid'] in selected_gids:continue
   ex=experience_by_candidate[r['gid'],approved.get(r['gid'],'')]
   reason='Sin identidad aprobada en LinkedIn' if r['gid'] not in approved else 'Sin experiencias guardadas' if not ex else 'Solo inicios anteriores al mes de grado' if all(e['postgraduation_status']=='excluded_start_before_graduation_month' for e in ex) else 'Otras restricciones de elegibilidad o cobertura'
   absence.append(dict(r,reason=reason))
  coverage[co]={'total':cell(cc,'coverage:total'),'combinations':partition(cc,'group',disclosure,'coverage:group',ensure=['Both sources','LinkedIn only','Comfama only','Neither source']),'absence_total':cell(absence,'absence:total'),'absence':partition(absence,'reason',disclosure,'coverage:absence')}
 write(out/'exact_filter_aggregates.json',exact);write(out/'longitudinal_exact.json',private_long)
 print('Protecting public filter lattice...',flush=True);disclosure.protect()
 def resolve(x):
  if isinstance(x,dict):
   if set(x)=={'$'}:return disclosure.release(x['$'])
   return {k:resolve(v) for k,v in x.items()}
  if isinstance(x,list):return [resolve(v) for v in x]
  return x
 public=resolve({'metadata':{'population':193,'approved_identities':157,'linkedin_graduates':114,'combined_graduates':144,'linkedin_experiences':323,'comfama_observations':57,'combined_observations':380,'postgraduate_marks':19,'practices':211,'privacy_k':K,'count_rounding':None,'size_minimum':K,'counts_approximate':False,'source_commit':'a0c276356c3852c2f6cf971ed4f52fa2d6aa69f7'},'professional_categories':PROFESSIONAL_AREAS,'cohorts':COHORTS,'views':views,'formation':formation,'coverage':coverage})
 for v in public['views'].values():
  for dim in ['economic_sector','professional_area']:
   safe={i['label'] for i in v[dim] if i['value'] is not None};v['details'][dim]={k:val for k,val in v['details'][dim].items() if k in safe}
  for sequence in v['longitudinal'].values():
   for r in sequence:
    if r['mean_status']=='available' and r['mean'] is None:r['mean_status']='suppressed'
 labels={'Promedio del tamaño empresarial','Otros grupos agrupados'};neighbors=defaultdict(set)
 def collect(x):
  if isinstance(x,dict):
   if 'label' in x:labels.add(x['label'])
   for value in x.values():collect(value)
  elif isinstance(x,list):
   peers={v['label'] for v in x if isinstance(v,dict) and 'label' in v}
   for label in peers:neighbors[label].update(peers-{label})
   for value in x:collect(value)
 collect(public)
 previous_payload=read(args.out) if args.out.exists() else {}
 previous=previous_payload.get('colors',{}) if previous_payload.get('color_registry_version')==3 else {}
 priority=PROFESSIONAL_AREAS+[i['label'] for i in public['views']['todas|0|0|0|1']['economic_sector']]+['No determinado','Otros grupos agrupados','Promedio del tamaño empresarial']
 public['colors']=categorical_colors(labels,previous,priority,neighbors);public['color_registry_version']=3
 write(args.out,public)
 write(out/'disclosure_report.json',{'k':K,'count_rounding':None,'minimum_size_statistic_graduates':K,'cells':len(disclosure.cells),'unique_signatures':len({c['sig'] for c in disclosure.cells}),'decisions':dict(disclosure.stats),'notes':['Named categories require two distinct graduates. Smaller groups and personal employers are pooled without dropping observations.','The four all-cohort source/postgraduate reference scenarios are released exactly when support permits.','Narrower filter comparisons in the same metric family are withheld when their difference isolates one known contributor.','If any component is suppressed, the whole distribution or size composition is withheld to avoid subtraction.','Reference tables are aggregate releases, not a guarantee of formal differential privacy.']})
 write(out/'disclosure_signatures_private.json',[{'records':c['sig'],'blocked':c['blocked'],'minimum':c['min'],'scalar':c['scalar'],'family':c['family'],'context':c['context'],'reference':c['reference']} for c in disclosure.cells])
 write(out/'disclosure_partitions_private.json',{'partitions':disclosure.partitions,'size_compositions':disclosure.compositions})
 print(json.dumps({'records':len(rows),'employers':len(registry),'size_missing':sum(a['size_estimate'] is None for a in registry.values()),'release_cells':dict(disclosure.stats),'output':str(args.out)},ensure_ascii=False))

if __name__=='__main__':main()
