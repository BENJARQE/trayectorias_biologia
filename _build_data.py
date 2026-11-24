"""Local aggregate generator. Never copy input workbooks or record-level outputs to the public repo.

Adapted from Jorge Arango's trayectorias_musica/_build_data.py (permission confirmed by requester).
Python + vanilla JS/SVG; no framework. All joins use source + record ID + employer_key.
"""
from pathlib import Path
from collections import Counter,defaultdict
from statistics import mean,median
from itertools import product
from datetime import date
import argparse,json,re,unicodedata,hashlib,math,csv
from openpyxl import load_workbook

K=5
COHORTS=['todas','2017–2019','2020–2021','2022–2023','2024–2025','2026 provisional']
UNKNOWN='No determinado'
# Original Music chart palettes, including its orange chart series, are retained.
PALETTE=[['#2563eb','#5b8dfa'],['#0f9d74','#2cc79a'],['#c026d3','#e167f0'],['#ea580c','#fb8a4a'],['#b45309','#e0a03a'],['#64748b','#94a3b8'],['#9ca3af','#6b7280']]

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
def rounded(n,step=5):return int(math.floor(n/step+0.5)*step)
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
 """Stable cell registry shared by every filter and both display units.

 Minimum five distinct graduates; complementary suppression; pairwise comparison
 guard across the entire released count lattice; rounding to base five. No raw
 signatures leave this process. Repeated queries reuse the same protected cell.
 """
 def __init__(self):self.cells=[];self.index={};self.stats=Counter();self.forced=set()
 def add(self,rows,scalar=None,minimum=K):
  # Contribution signature includes record presence, even when totals coincide.
  sig=tuple(sorted((r['uid'],r['gid']) for r in rows));key=(sig,scalar,minimum)
  if key in self.index:return self.index[key]
  i=len(self.cells);self.index[key]=i
  people=frozenset(g for _,g in sig if g)
  self.cells.append({'sig':sig,'people':people,'n':len(sig),'scalar':scalar,'min':minimum,'blocked':bool(sig) and len(people)<minimum})
  return i
 def protect(self):
  # Catch narrow complements between cells, including aggregate/cohort,
  # source switches, subgroup clicks, postgraduate and Others exclusions.
  # Prefer the unfiltered reference views registered first. If a narrower view
  # would disclose a small difference, suppress that view, not its safe parent.
  for c in self.cells:
   if c['sig'] in self.forced:c['blocked']=True
  signatures=list(dict.fromkeys(c['sig'] for c in self.cells if not c['blocked']))
  record_people={u:g for sig in signatures for u,g in sig};rs=[frozenset(u for u,g in sig) for sig in signatures]
  blocked=set();accepted=[]
  for i,a in enumerate(rs):
   for j in accepted:
    b=rs[j];difference=a^b
    if not difference:continue
    people={record_people[u] for u in difference if record_people[u]}
    if len(people)<K:
     blocked.add(signatures[i]);break
   if signatures[i] not in blocked:accepted.append(i)
  for c in self.cells:
   if c['sig'] in blocked:c['blocked']=True
   if not c['sig']:c['blocked']=False
   self.stats['suppressed' if c['blocked'] else 'released']+=1
 def complementary(self,groups):
  hidden={r['gid'] for rs in groups.values() if len({r['gid'] for r in rs if r['gid']})<K for r in rs if r['gid']}
  visible=[rs for rs in groups.values() if len({r['gid'] for r in rs if r['gid']})>=K]
  if 0<len(hidden)<K and visible:
   extra=min(visible,key=lambda rs:len(rs))
   self.forced.add(tuple(sorted((r['uid'],r['gid']) for r in extra)))
 def release(self,i):
  c=self.cells[i]
  if c['blocked']:return None
  if c['scalar'] is not None:return rounded(c['scalar'],100) if c['scalar']>=100 else round(c['scalar'],1)
  return {'exp':rounded(c['n']),'grad':rounded(len(c['people']))}

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
   rows.append(dict(r,**{k:v for k,v in a.items() if k not in ['audit_original','size_audit_original']},uid=src+':'+eid,gid=r['graduate_id'],source=src,cohort=cohort(gmap[r['graduate_id']]['graduation_date']),professional_area=area or UNKNOWN,title_normalized=clean(r['title']) if src=='LinkedIn' else UNKNOWN,country=r.get('country_normalized') or UNKNOWN,city=r.get('city_normalized') or UNKNOWN,postgraduate=r.get('exclude_postgraduate_scenario',False),position=r.get('position_order'),area_basis='Cargo documentado' if src=='LinkedIn' else 'Proxy de actividad del empleador; cargo no observado'))
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
   r.update({k:a.get(k) for k in extra});r['professional_area']=r['sector'];r['title_normalized']=clean(r['title'])
  sheet.update(tab(rs,sheet['headers']+[h for h in extra if h not in sheet['headers']]))
 review['sheets'].append(dict(name='Employer audit',**tab([{k:(json.dumps(v,ensure_ascii=False) if isinstance(v,dict) else v) for k,v in a.items()} for a in registry.values()])))
 write(out/'review_enriched.json',review)
 with (out/'employer_sizes_pending.csv').open('w',encoding='utf-8-sig',newline='') as f:
  fields=['employer_key','employer_name','size_estimate','size_method','size_evidence_type','size_confidence','size_low','size_scope','size_reference_date','size_evidence','size_warning','size_pending'];w=csv.DictWriter(f,fieldnames=fields,extrasaction='ignore');w.writeheader();w.writerows(registry.values())
 # Normalize title typography only. No semantic translations or area reassignment.
 title_labels={}
 for r in li:title_labels.setdefault(token(r['title_normalized']),r['title_normalized'])
 for r in li:r['title_normalized']=title_labels[token(r['title_normalized'])]
 disclosure=Disclosure();views={};exact={};private_long=[]
 def cell(rs):return {'$':disclosure.add(rs)}
 def scalar(rs,v):return {'$':disclosure.add(rs,v,10)} if len({r['gid'] for r in rs})>=10 else None
 def distribution(rs,field,unknown=True):
  groups=defaultdict(list)
  for r in rs:
   label=r[field] or UNKNOWN
   if not unknown and is_unknown(label):continue
   if field=='employer_name' and (r['employer_personal'] or token(label)=='profesional independiente'):label='Empleadores reservados o independientes'
   groups[label].append(r)
  # Use the same label universe for all queries; do not publish rare category names.
  disclosure.complementary(groups)
  return [{'label':label,'value':cell(v)} for label,v in sorted(groups.items()) if len({r['gid'] for r in v})>=K or label==UNKNOWN]
 def subset(co,comf,pg,other):return [r for r in rows if (co=='todas' or r['cohort']==co) and (comf or r['source']=='LinkedIn') and not(pg and r['postgraduate']) and not(other and r['professional_area']=='Otros')]
 for co,comf,pg,other,nd in product(COHORTS,[0,1],[0,1],[0,1],[0,1]):
  key='|'.join(map(str,[co,comf,pg,other,nd]));rs=subset(co,comf,pg,other);linkedin=[r for r in rs if r['source']=='LinkedIn']
  v={'total':cell(rs),'linkedin':cell(linkedin),'comfama':cell([r for r in rs if r['source']=='Comfama']),'economic_sector':distribution(rs,'economic_sector',nd),'professional_area':distribution(rs,'professional_area',nd),'employers':distribution(rs,'employer_name',nd),'titles':distribution(linkedin,'title_normalized',nd),'countries':distribution(rs,'country',nd),'cities':distribution(rs,'city',nd),'details':{},'longitudinal':{}}
  for dim in ['economic_sector','professional_area']:
   v['details'][dim]={}
   for item in v[dim]:
    group=[r for r in rs if r[dim]==item['label']];v['details'][dim][item['label']]={'employers':distribution(group,'employer_name',nd),'titles':distribution([r for r in group if r['source']=='LinkedIn'],'title_normalized',nd)}
  # First and latest refer to immutable original selected positions.
  for which,positions in [('first',{g:1 for g in counts}),('last',dict(counts))]:
   selected=[r for r in linkedin if r['position']==positions[r['gid']]]
   v[which]={'total':cell(selected),'sectors':distribution(selected,'economic_sector',nd),'areas':distribution(selected,'professional_area',nd)}
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
  v['international']={'known':cell(known),'foreign':cell(foreign),'unknown':cell([r for r in rs if r['country']==UNKNOWN])}
  for balanced,low in product([0,1],[0,1]):
   seq=[]
   for pos in range(1,7):
    eligible=[r for r in originals if r['position']==pos and (not balanced or r['gid'] in fixed)]
    selected=[r for r in linkedin if r['position']==pos and (not balanced or r['gid'] in fixed)]
    sized=[r for r in selected if r['size_estimate'] is not None and not(low and r['size_low'])]
    means=mean(r['size_estimate'] for r in sized) if sized else None;med=median(r['size_estimate'] for r in sized) if sized else None
    seq.append({'position':pos,'eligible':cell(eligible),'retained':cell(selected),'usable':cell(sized),'mean':scalar(sized,means) if sized else None,'median':scalar(sized,med) if sized else None,'sectors':distribution(selected,'economic_sector',nd),'areas':distribution(selected,'professional_area',nd)})
    if not comf and not nd:private_long.append(dict(cohort=co,exclude_postgraduate=pg,exclude_other=other,fixed_sample=balanced,exclude_low_size=low,position=pos,eligible=len(eligible),retained=len(selected),usable=len(sized),coverage=len(sized)/len(eligible) if eligible else None,mean=means,median=med))
   v['longitudinal'][f'{balanced}|{low}']=seq
  views[key]=v
  exact[key]={'linkedin':len(linkedin),'comfama':len(rs)-len(linkedin),'observations':len(rs),'graduates':len({r['gid'] for r in rs}),'sectors':dict(Counter(r['economic_sector'] for r in rs)),'areas':dict(Counter(r['professional_area'] for r in rs))}
 # Practices and M0 obey only compatible filters: cohort; PG for continuity.
 pt=unpack(practice['tables']['Practice Trace']);pairs=unpack(practice['tables']['Continuity']);m0=unpack(practice['tables']['M0 Trace']);formation={}
 for co,pg in product(COHORTS,[0,1]):
  chosen=[r for r in pt if r['Included in aggregates'] and (co=='todas' or r['Cohort']==co)]
  pr=[dict(uid=r['Practice record id'],gid=r['Graduate id'] or None,entity=r['Entity canonical'],country=r['Country reviewed'] or r['Country group'] or UNKNOWN,city=r['City reviewed'] or r['City (original)'] or UNKNOWN,scope=r['Geographic scope']) for r in chosen]
  cv=[]
  for scope in ['All entities','EAFIT','Other entities']:
   group=[r for r in chosen if scope=='All entities' or (('eafit' in token(r['Entity canonical']))==(scope=='EAFIT'))]
   people={r['Graduate id'] for r in group if r['Graduate id']};in_final={r['gid'] for r in li if not(pg and r['postgraduate'])};accepted=[r for r in pairs if r['Accepted all'] and not(pg and r['Exclude postgraduate scenario']) and (co=='todas' or r['Cohort']==co) and (scope=='All entities' or r['Entity scope']==scope)]
   def person_rows(ids,prefix):return [dict(uid='person:'+g,gid=g) for g in ids]
   cv.append({'scope':scope,'practice_graduates':cell(person_rows(people,'practice')),'jointly_observed':cell(person_rows(people&in_final,'joint')),'graduates':cell(person_rows({r['Graduate id'] for r in accepted},'continuity')),'matches':cell([dict(uid=r['Employment id'],gid=r['Graduate id']) for r in accepted])})
  chosen_m0=[r for r in m0 if r['Selected latest for graduate'] and (co=='todas' or r['Cohort']==co)]
  mr=[dict(uid='m0:'+str(r['Graduate ID']),gid=str(r['Graduate ID']),paid=r['Paid work status'],offer=r['Offer status'],accepted=r['Acceptance status'],working=r['Practice-entity work status'],timing=r['Response timing vs graduation']) for r in chosen_m0]
  # Apply the same all-cohort offer coverage rule already closed in the source:
  # early-cohort form did not establish offer-item comparability.
  for rr in mr:
   if cohort(gmap[rr['gid']]['graduation_date'])==COHORTS[1]:rr['offer']='Missing';rr['accepted']='Missing';rr['working']='Missing'
  formation[f'{co}|{pg}']={'total':cell(pr),'entities':distribution_generic(pr,'entity',disclosure),'countries':distribution_generic(pr,'country',disclosure),'cities':distribution_generic(pr,'city',disclosure),'geography':distribution_generic(pr,'scope',disclosure),'continuity':cv,'m0':{'respondents':cell(mr),'paid':distribution_generic(mr,'paid',disclosure),'offer':distribution_generic(mr,'offer',disclosure),'accepted':distribution_generic(mr,'accepted',disclosure),'working':distribution_generic(mr,'working',disclosure),'timing':distribution_generic(mr,'timing',disclosure)}}
 # Coverage does not change when a recorded episode is excluded by a chart filter.
 coverage={}
 cov=unpack(com['tables']['Coverage'])
 for co in COHORTS:
  cc=[dict(uid='coverage:'+r['graduate_id'],gid=r['graduate_id'],group=r['coverage_group']) for r in cov if co=='todas' or cohort(gmap[r['graduate_id']]['graduation_date'])==co]
  coverage[co]=distribution_generic(cc,'group',disclosure)
 write(out/'exact_filter_aggregates.json',exact);write(out/'longitudinal_exact.json',private_long)
 print('Protecting public filter lattice...',flush=True);disclosure.protect()
 def resolve(x):
  if isinstance(x,dict):
   if set(x)=={'$'}:return disclosure.release(x['$'])
   return {k:resolve(v) for k,v in x.items()}
  if isinstance(x,list):return [resolve(v) for v in x]
  return x
 public=resolve({'metadata':{'date':'2026-11-24','population':193,'approved_identities':157,'linkedin_graduates':114,'combined_graduates':144,'linkedin_experiences':323,'comfama_episodes':57,'combined_observations':380,'postgraduate_marks':19,'practices':211,'privacy_k':K,'count_rounding':5,'size_minimum':10,'counts_approximate':True,'source_commit':'a0c276356c3852c2f6cf971ed4f52fa2d6aa69f7'},'cohorts':COHORTS,'palette':PALETTE,'views':views,'formation':formation,'coverage':coverage})
 # Remove labels and empty drilldown objects if disclosure blocked every count.
 def scrub(x):
  if isinstance(x,list):return [scrub(v) for v in x if not(isinstance(v,dict) and 'value' in v and v['value'] is None)]
  if isinstance(x,dict):return {k:scrub(v) for k,v in x.items()}
  return x
 public=scrub(public)
 for v in public['views'].values():
  for dim in ['economic_sector','professional_area']:
   safe={i['label'] for i in v[dim]};v['details'][dim]={k:val for k,val in v['details'][dim].items() if k in safe}
 write(args.out,public)
 write(out/'disclosure_report.json',{'k':K,'count_rounding':5,'minimum_size_statistic_graduates':10,'cells':len(disclosure.cells),'unique_signatures':len({c['sig'] for c in disclosure.cells}),'decisions':dict(disclosure.stats),'notes':['Suppression is shared by every occurrence of the same signature, including unit switches.','Pairs differing in fewer than five contributors are suppressed across the complete precomputed query lattice.','No small category labels or person employer names are serialized.','Counts are rounded to base five; percentages are derived only from released rounded values.','Exact institutional controls are explicitly authorized baseline totals.','Statistical disclosure controls reduce risk; they do not establish formal differential privacy.']})
 write(out/'disclosure_signatures_private.json',[{'records':c['sig'],'blocked':c['blocked'],'minimum':c['min'],'scalar':c['scalar']} for c in disclosure.cells])
 print(json.dumps({'records':len(rows),'employers':len(registry),'size_missing':sum(a['size_estimate'] is None for a in registry.values()),'release_cells':dict(disclosure.stats),'output':str(args.out)},ensure_ascii=False))

def distribution_generic(rows,field,disclosure):
 groups=defaultdict(list)
 for r in rows:groups[r[field] or UNKNOWN].append(r)
 disclosure.complementary(groups)
 return [{'label':label,'value':{'$':disclosure.add(rs)}} for label,rs in sorted(groups.items()) if len({r['gid'] for r in rs if r['gid']})>=K or label==UNKNOWN]
if __name__=='__main__':main()
