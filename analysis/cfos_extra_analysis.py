"""Código del análisis final de c-Fos; véase analysis/README.md para su alcance."""
import argparse
from pathlib import Path
import sys,itertools,numpy as np,pandas as pd
from scipy import stats
import statsmodels.api as sm
from statsmodels.stats.multitest import multipletests
from read_xlsx_xml import XLSX


def run(input_dir, output_dir):
    ROOT=str(Path(input_dir).expanduser().resolve())
    OUT=str(Path(output_dir).expanduser().resolve())+'/'
    a=pd.read_csv(OUT+'datos_animal.csv');q=pd.read_csv(OUT+'qc_stacks.csv');base=pd.read_csv(OUT+'polaridad_animal.csv')
    X=XLSX(ROOT+'/resultados_cfos_2D_3D.xlsx');sec=X.read('Sectores_por_animal');sec['Animal_ID']=sec.Experimento.astype(str)+'::'+sec.Animal.astype(str)
    for c in ['Celulas_cFos','Volumen_um3','Densidad_celulas_mm3']:sec[c]=pd.to_numeric(sec[c],errors='coerce')

    def dm(d,tr,co,ap=False):
     x=d[d.Grupo.isin([tr,co])].copy();keep=[e for e,z in x.groupby('Experimento') if set(z.Grupo)=={tr,co}];x=x[x.Experimento.isin(keep)].sort_values(['Experimento','Animal_ID']).reset_index(drop=True);T=(x.Grupo==tr).astype(float).to_numpy();ex=pd.get_dummies(x.Experimento,drop_first=True,dtype=float);cols=[np.ones(len(x)),T];
     if ex.shape[1]:cols.extend([ex[c].to_numpy() for c in ex])
     if ap:
      z=x.Bregma_medio.to_numpy();z=(z-z.mean())/(z.std(ddof=0) or 1);cols.append(z)
     return x,T,np.column_stack(cols)
    def perms(x,T):
     blocks=[]
     for e,z in x.groupby('Experimento',sort=False):
      idx=z.index.to_numpy();k=int(T[idx].sum());blocks.append((idx,list(itertools.combinations(idx,k))))
     for part in itertools.product(*[b[1] for b in blocks]):
      t=np.zeros(len(x));
      for ind in part:t[list(ind)]=1
      yield t
    def fit(y,X,hc=True):
     f=sm.OLS(y,X).fit(cov_type='HC3') if hc else sm.OLS(y,X).fit();return f.params[1],f.bse[1],f.tvalues[1]
    def tst(d,metric,tr,co,ap=False,exp=True,hc=True):
     x,T,X=dm(d,tr,co,ap);y=x[metric].to_numpy(float);b,se,to=fit(y,X,hc);vv=[]
     for tp in perms(x,T):
      xp=X.copy();xp[:,1]=tp;vv.append(abs(fit(y,xp,hc)[2]))
     p=np.mean(np.array(vv)>=abs(to)-1e-12);df=max(len(y)-np.linalg.matrix_rank(X),1);cr=stats.t.ppf(.975,df)
     return dict(Region=x.Region.iloc[0],Tratamiento=tr,Control=co,n_T=int(T.sum()),n_C=int(len(T)-T.sum()),Efecto=np.exp(b) if exp else b,Efecto_IC95_inf=np.exp(b-cr*se) if exp else b-cr*se,Efecto_IC95_sup=np.exp(b+cr*se) if exp else b+cr*se,p_exacta=p,Permutaciones=len(vv))
    # Classic t exact permutation.
    rows=[]
    for r in ['BLAa','BLAp']:
     for tr in ['SCOP-HB','LiCl']:rows.append(tst(a[a.Region==r],'Log_densidad',tr,'Vehículo',hc=False))
    z=pd.DataFrame(rows);z['p_Holm']=multipletests(z.p_exacta,method='holm')[1];z.to_csv(OUT+'densidad_tclasica.csv',index=False)
    # AP adjustment global.
    rows=[]
    for r in ['BLAa','BLAp']:
     for tr in ['SCOP-HB','LiCl']:rows.append(tst(a[a.Region==r],'Log_densidad',tr,'Vehículo',ap=True))
    z=pd.DataFrame(rows);z['p_Holm']=multipletests(z.p_exacta,method='holm')[1];z.to_csv(OUT+'densidad_ajuste_AP.csv',index=False)
    # Exclude high outside ROI warning stacks and recompute animal densities.
    warn=q.Advertencias.notna();qw=q[~warn].copy();ag=qw.groupby(['Experimento','Animal','Animal_ID','Grupo','Region']).agg(Celulas_cFos=('Objetos_validos','sum'),Volumen_um3=('Volumen_um3','sum'),N_stacks=('Stack_ID','count')).reset_index();ag['Densidad_celulas_mm3']=ag.Celulas_cFos/ag.Volumen_um3*1e9;ag['Log_densidad']=np.log(ag.Densidad_celulas_mm3);ag=ag.merge(a[['Animal_ID','Region','Bregma_medio']],on=['Animal_ID','Region'])
    rows=[]
    for r in ['BLAa','BLAp']:
     for tr in ['SCOP-HB','LiCl']:rows.append(tst(ag[ag.Region==r],'Log_densidad',tr,'Vehículo'))
    z=pd.DataFrame(rows);z['p_Holm']=multipletests(z.p_exacta,method='holm')[1];z.to_csv(OUT+'densidad_sin_warnings.csv',index=False)
    # Only animal-region with 2 stacks.
    a2=a[a.N_stacks>=2]
    rows=[]
    for r in ['BLAa','BLAp']:
     for tr in ['SCOP-HB','LiCl']:rows.append(tst(a2[a2.Region==r],'Log_densidad',tr,'Vehículo'))
    z=pd.DataFrame(rows);z['p_Holm']=multipletests(z.p_exacta,method='holm')[1];z.to_csv(OUT+'densidad_dos_stacks.csv',index=False)
    # Leave-one-out effect ranges, no p.
    loo=[]
    for r in ['BLAa','BLAp']:
     for tr in ['SCOP-HB','LiCl']:
      d=a[(a.Region==r)&a.Grupo.isin([tr,'Vehículo'])]
      for aid in d.Animal_ID.unique():
       dd=d[d.Animal_ID!=aid]
       try:
        x,T,X=dm(dd,tr,'Vehículo',False);b,se,tv=fit(x.Log_densidad.to_numpy(),X,True);loo.append(dict(Region=r,Tratamiento=tr,Animal_omitido=aid,Efecto=np.exp(b)))
       except:pass
    pd.DataFrame(loo).to_csv(OUT+'influencia_animal.csv',index=False)
    # Half absolute density tests. Use pseudocount 0.5 to handle zeros.
    for side in ['Lateral','Medial','Dorsal','Ventral']:
     base['Log_abs_'+side]=np.log((base['Celulas_'+side]+.5)/base['Volumen_'+side]*1e9)
    rows=[]
    for r in ['BLAa','BLAp']:
     for tr in ['SCOP-HB','LiCl']:
      for side in ['Lateral','Medial','Dorsal','Ventral']:
       zz=tst(base[base.Region==r],'Log_abs_'+side,tr,'Vehículo',ap=True);zz['Mitad']=side;rows.append(zz)
    z=pd.DataFrame(rows);z['p_Holm']=multipletests(z.p_exacta,method='holm')[1];z.to_csv(OUT+'densidad_mitades.csv',index=False)
    # Quadrant animal data and tests.
    qtot=a[['Animal_ID','Region','Experimento','Animal','Grupo','Bregma_medio']]
    quad=sec.merge(qtot,on=['Animal_ID','Region','Experimento','Animal','Grupo'],how='left');quad['Log_dens']=np.log((quad.Celulas_cFos+.5)/quad.Volumen_um3*1e9);quad.to_csv(OUT+'cuadrantes_animal.csv',index=False)
    rows=[]
    for r in ['BLAa','BLAp']:
     for tr in ['SCOP-HB','LiCl']:
      for sector in sorted(quad.Sector.unique()):
       d=quad[(quad.Region==r)&(quad.Sector==sector)&(quad.Volumen_um3>0)]
       try:zz=tst(d,'Log_dens',tr,'Vehículo',ap=True);zz['Cuadrante']=sector;rows.append(zz)
       except:pass
    z=pd.DataFrame(rows);z['p_Holm']=multipletests(z.p_exacta,method='holm')[1];z.to_csv(OUT+'densidad_cuadrantes.csv',index=False)
    print('classic\n',pd.read_csv(OUT+'densidad_tclasica.csv').to_string(index=False));print('\nAP\n',pd.read_csv(OUT+'densidad_ajuste_AP.csv').to_string(index=False));print('\nWARN\n',pd.read_csv(OUT+'densidad_sin_warnings.csv').to_string(index=False));print('\n2stacks\n',pd.read_csv(OUT+'densidad_dos_stacks.csv').to_string(index=False));print('\nHALF min pHolm',pd.read_csv(OUT+'densidad_mitades.csv').sort_values('p_Holm').head(8).to_string(index=False));print('\nQUAD min pHolm',pd.read_csv(OUT+'densidad_cuadrantes.csv').sort_values('p_Holm').head(8).to_string(index=False))


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', required=True, type=Path, help='Carpeta de la corrida original')
    parser.add_argument('--output', required=True, type=Path, help='Resultados existentes de cfos_final_analysis.py')
    args=parser.parse_args()
    run(args.input, args.output)


if __name__ == "__main__":
    main()
