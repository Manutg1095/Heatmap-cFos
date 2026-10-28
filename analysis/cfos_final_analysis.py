"""Código del análisis final de c-Fos; véase analysis/README.md para su alcance."""
import argparse
from pathlib import Path
import sys, math, itertools, json, hashlib, zipfile
import numpy as np, pandas as pd
from scipy import stats
import statsmodels.api as sm
from statsmodels.stats.multitest import multipletests
from read_xlsx_xml import XLSX


def run(input_dir, output_dir):
    ROOT=Path(input_dir).expanduser().resolve()
    OUT=Path(output_dir).expanduser().resolve()
    OUT.mkdir(parents=True, exist_ok=True)
    X=XLSX(ROOT/'resultados_cfos_2D_3D.xlsx')
    S={s:X.read(s) for s in ['Version','Convenciones_espaciales','Stacks_registrados','Stacks_no_ejecutados','QC_y_metricas_stack','Analisis_principal_animal','Sectores_por_animal','Bins_2D_stack','Celulas_3D','Resumen_3D_grupo']}
    XC='X_ML_norm_0_medial_1_lateral'; YC='Y_DV_norm_0_ventral_1_dorsal'; ACX='Centroide_X_ML_norm_0_medial_1_lateral'; ACY='Centroide_Y_DV_norm_0_ventral_1_dorsal'
    for d in S.values():
        if isinstance(d,pd.DataFrame):
            for c in d.columns:
                if c in ['Bregma_mm','Area_um2','Volumen_um3','Objetos_XYZ','Objetos_incluidos_filtro','Objetos_fuera_ROI','Objetos_validos','Celulas_cFos','Densidad_celulas_mm3','N_stacks',XC,YC,ACX,ACY,'Centroide_AP_mm','AP_mm','Bin_X','Bin_Y']:
                    d[c]=pd.to_numeric(d[c],errors='coerce')
    q=S['QC_y_metricas_stack'].copy(); a=S['Analisis_principal_animal'].copy(); sec=S['Sectores_por_animal'].copy(); bins=S['Bins_2D_stack'].copy(); cells=S['Celulas_3D'].copy()
    for d in [q,a,sec,bins,cells]:
        d['Animal_ID']=d['Experimento'].astype(str)+'::'+d['Animal'].astype(str)
    # Animal-level bregma weighted by volume.
    bq=q.copy(); bq['bv']=bq.Bregma_mm*bq.Volumen_um3
    bm=bq.groupby(['Animal_ID','Region']).agg(Bregma_medio=('bv','sum'),Vol=('Volumen_um3','sum')).reset_index();bm.Bregma_medio/=bm.Vol
    a=a.merge(bm[['Animal_ID','Region','Bregma_medio']],on=['Animal_ID','Region'],how='left')
    a['Log_densidad']=np.log(a.Densidad_celulas_mm3)
    # Core QC summaries.
    summary=[]
    def check(k,v):summary.append((k,v))
    check('Animales',a.Animal_ID.nunique()); check('Filas animal-region',len(a));check('Stacks ejecutados',len(q)); check('Objetos XYZ',int(q.Objetos_XYZ.sum()));check('Incluidos filtro',int(q.Objetos_incluidos_filtro.sum()));check('Fuera ROI',int(q.Objetos_fuera_ROI.sum()));check('Objetos válidos',int(q.Objetos_validos.sum()));check('Volumen um3',float(q.Volumen_um3.sum()))
    # Reconcile animal aggregates.
    agg=q.groupby(['Animal_ID','Region']).agg(Celulas=('Objetos_validos','sum'),Volumen=('Volumen_um3','sum'),Stacks=('Stack_ID','count')).reset_index();mer=a.merge(agg,on=['Animal_ID','Region']);check('Máx delta células animal',int(np.max(np.abs(mer.Celulas_cFos-mer.Celulas))));check('Máx delta volumen animal',float(np.max(np.abs(mer.Volumen_um3-mer.Volumen))));check('Máx delta stacks animal',int(np.max(np.abs(mer.N_stacks-mer.Stacks))))
    # Conteo workbook reconciliation.
    CX=XLSX(ROOT/'conteo_objetos.xlsx'); cont=CX.read('Conteo_por_stack')
    for c in ['Objetos_totales_XYZ','Objetos_post_filtro','Objetos_fuera_ROI','Objetos_validos_ROI']:
        cont[c]=pd.to_numeric(cont[c],errors='coerce')
    cm=cont.merge(q[['Stack_ID','Objetos_XYZ','Objetos_incluidos_filtro','Objetos_fuera_ROI','Objetos_validos']],on='Stack_ID',how='left')
    cm['Delta_postfiltro']=cm.Objetos_post_filtro-cm.Objetos_incluidos_filtro
    cm['Delta_validos']=cm.Objetos_validos_ROI-cm.Objetos_validos
    cm[cm.Delta_postfiltro.fillna(0)!=0].to_csv(OUT/'discrepancias_postfiltro.csv',index=False)
    # Design.
    design=a[a.Region=='BLAa'].groupby(['Experimento','Grupo']).Animal_ID.nunique().unstack(fill_value=0).reset_index();design.to_csv(OUT/'diseno.csv',index=False)
    # Descriptives.
    desc=a.groupby(['Region','Grupo']).Densidad_celulas_mm3.agg(n='count',Media='mean',DE='std',Mediana='median',Q1=lambda x:x.quantile(.25),Q3=lambda x:x.quantile(.75)).reset_index();desc.to_csv(OUT/'descriptivos.csv',index=False)
    counts=q.groupby(['Region','Grupo']).agg(N_stacks=('Stack_ID','count'),Validos=('Objetos_validos','sum'),Volumen_um3=('Volumen_um3','sum')).reset_index(); counts=counts.merge(a.groupby(['Region','Grupo']).Animal_ID.nunique().rename('N_animales').reset_index(),on=['Region','Grupo']);counts.to_csv(OUT/'conteo_grupo.csv',index=False)
    # Exact stratified permutation regression.
    def design_matrix(d,treated,control,with_ap=False):
        x=d[d.Grupo.isin([treated,control])].copy()
        # retain experiments containing both labels
        keep=[e for e,z in x.groupby('Experimento') if set(z.Grupo)=={treated,control}]
        x=x[x.Experimento.isin(keep)].copy().sort_values(['Experimento','Animal_ID']).reset_index(drop=True)
        T=(x.Grupo==treated).astype(float).to_numpy()
        # intercept + experiment dummies (drop first) + optional AP standardized + T
        ex=pd.get_dummies(x.Experimento,drop_first=True,dtype=float)
        cols=[np.ones(len(x)),T]
        if ex.shape[1]:cols.extend([ex[c].to_numpy() for c in ex])
        if with_ap:
            z=x.Bregma_medio.to_numpy(float); z=(z-z.mean())/(z.std(ddof=0) or 1);cols.append(z)
        X=np.column_stack(cols)
        return x,T,X

    def ols_hc3(y,X):
        fit=sm.OLS(y,X).fit(cov_type='HC3')
        return float(fit.params[1]),float(fit.bse[1]),float(fit.tvalues[1]),fit

    def strat_perms(x,T):
        blocks=[]
        for e,z in x.groupby('Experimento',sort=False):
            idx=z.index.to_numpy(); k=int(T[idx].sum()); blocks.append((idx,list(itertools.combinations(idx,k))))
        for parts in itertools.product(*[b[1] for b in blocks]):
            tp=np.zeros(len(x))
            for inds in parts: tp[list(inds)]=1
            yield tp

    def test_log_metric(d,metric,treated,control,with_ap=False):
        x,T,X=design_matrix(d,treated,control,with_ap); y=x[metric].to_numpy(float)
        beta,se,tobs,fit=ols_hc3(y,X)
        vals=[]
        for tp in strat_perms(x,T):
            Xp=X.copy();Xp[:,1]=tp
            try:vals.append(abs(ols_hc3(y,Xp)[2]))
            except:pass
        p=float(np.mean(np.array(vals)>=abs(tobs)-1e-12))
        df=max(len(y)-np.linalg.matrix_rank(X),1);crit=stats.t.ppf(.975,df)
        return dict(Region=x.Region.iloc[0],Tratamiento=treated,Control=control,n_T=int(T.sum()),n_C=int(len(T)-T.sum()),Efecto=float(np.exp(beta)),Efecto_IC95_inf=float(np.exp(beta-crit*se)),Efecto_IC95_sup=float(np.exp(beta+crit*se)),p_exacta=p,Permutaciones=len(vals),t_HC3=tobs)
    # Global density family.
    rows=[]
    for reg in ['BLAa','BLAp']:
        dr=a[a.Region==reg]
        for tr in ['SCOP-HB','LiCl']:rows.append(test_log_metric(dr,'Log_densidad',tr,'Vehículo',False))
    glob=pd.DataFrame(rows);glob['p_Holm']=multipletests(glob.p_exacta,method='holm')[1];glob.to_csv(OUT/'densidad_principal.csv',index=False)
    # Direct active comparison.
    rows=[]
    for reg in ['BLAa','BLAp']:rows.append(test_log_metric(a[a.Region==reg],'Log_densidad','LiCl','SCOP-HB',False))
    act=pd.DataFrame(rows);act['p_Holm']=multipletests(act.p_exacta,method='holm')[1];act.to_csv(OUT/'comparacion_activos.csv',index=False)
    # Directional halves from sectors.
    wide=sec.pivot_table(index=['Experimento','Animal','Animal_ID','Grupo','Region'],columns='Sector',values=['Celulas_cFos','Volumen_um3'],aggfunc='sum').reset_index()
    wide.columns=['_'.join([str(x) for x in c if str(x)!='']) if isinstance(c,tuple) else c for c in wide.columns]
    # easier aggregate original sectors directly.
    base=a[['Experimento','Animal','Animal_ID','Grupo','Region','Celulas_cFos','Volumen_um3','Bregma_medio',ACX,ACY]].copy().rename(columns={ACX:'Centroide_X',ACY:'Centroide_Y'})
    parts=[]
    for _,r in sec.iterrows():
        lat='Lateral' in r.Sector; dor='Dorsal' in r.Sector
        parts.append((r.Animal_ID,r.Region,r.Experimento,r.Animal,r.Grupo,'Lateral' if lat else 'Medial','Dorsal' if dor else 'Ventral',r.Celulas_cFos,r.Volumen_um3))
    sd=pd.DataFrame(parts,columns=['Animal_ID','Region','Experimento','Animal','Grupo','ML','DV','Cells','Vol'])
    for axis in ['ML','DV']:
        vv=sd.groupby(['Animal_ID','Region',axis]).agg(Cells=('Cells','sum'),Vol=('Vol','sum')).reset_index()
        cw=vv.pivot(index=['Animal_ID','Region'],columns=axis,values='Cells');vw=vv.pivot(index=['Animal_ID','Region'],columns=axis,values='Vol')
        for side in cw.columns:
            base=base.merge(cw[[side]].rename(columns={side:'Celulas_'+side}).reset_index(),on=['Animal_ID','Region'],how='left')
            base=base.merge(vw[[side]].rename(columns={side:'Volumen_'+side}).reset_index(),on=['Animal_ID','Region'],how='left')
    for side in ['Lateral','Medial','Dorsal','Ventral']:
        base['Densidad_'+side]=base['Celulas_'+side]/base['Volumen_'+side]*1e9
        base['Log_densidad_'+side]=np.log((base['Celulas_'+side]+0.5)/base['Volumen_'+side]*1e9)
    base['Log_RR_LM']=base.Log_densidad_Lateral-base.Log_densidad_Medial
    base['Log_RR_DV']=base.Log_densidad_Dorsal-base.Log_densidad_Ventral
    base.to_csv(OUT/'polaridad_animal.csv',index=False)
    # Direction tests.
    rows=[]
    for reg in ['BLAa','BLAp']:
      for tr in ['SCOP-HB','LiCl']:
        for axis,metric in [('ML','Log_RR_LM'),('DV','Log_RR_DV')]:
          z=test_log_metric(base[base.Region==reg],metric,tr,'Vehículo',True);z['Eje']=axis;rows.append(z)
    dir2=pd.DataFrame(rows);dir2['p_Holm']=multipletests(dir2.p_exacta,method='holm')[1];dir2.to_csv(OUT/'direccion_2D.csv',index=False)
    # Centroid tests; effect additive rather than exponential.
    def test_linear(d,metric,treated,control,with_ap=True):
        x,T,X=design_matrix(d,treated,control,with_ap); y=x[metric].to_numpy(float);beta,se,tobs,fit=ols_hc3(y,X);vals=[]
        for tp in strat_perms(x,T):
            Xp=X.copy();Xp[:,1]=tp
            try:vals.append(abs(ols_hc3(y,Xp)[2]))
            except:pass
        p=float(np.mean(np.array(vals)>=abs(tobs)-1e-12));df=max(len(y)-np.linalg.matrix_rank(X),1);crit=stats.t.ppf(.975,df)
        return dict(Region=x.Region.iloc[0],Tratamiento=treated,Control=control,n_T=int(T.sum()),n_C=int(len(T)-T.sum()),Efecto=beta,Efecto_IC95_inf=beta-crit*se,Efecto_IC95_sup=beta+crit*se,p_exacta=p,Permutaciones=len(vals))
    rows=[]
    for reg in ['BLAa','BLAp']:
     for tr in ['SCOP-HB','LiCl']:
      for axis,metric in [('ML','Centroide_X'),('DV','Centroide_Y')]:
       z=test_linear(base[base.Region==reg],metric,tr,'Vehículo',True);z['Eje']=axis;rows.append(z)
    cent=pd.DataFrame(rows);cent['p_Holm']=multipletests(cent.p_exacta,method='holm')[1];cent.to_csv(OUT/'centroides_2D.csv',index=False)
    # Descriptive profiles.
    profiles=[]
    for (g,r),d in base.groupby(['Grupo','Region']):
     for axx,n,dn in [('ML','Lateral','Medial'),('DV','Dorsal','Ventral')]:
      ratio=d['Densidad_'+n]/d['Densidad_'+dn]
      profiles.append(dict(Grupo=g,Region=r,Eje=axx,n=len(d),Razon_mediana=ratio.median(),Q1=ratio.quantile(.25),Q3=ratio.quantile(.75),Animales_mayor_1=int((ratio>1).sum()),Fraccion_objetos_media=(d['Celulas_'+n]/d.Celulas_cFos).mean(),Fraccion_volumen_media=(d['Volumen_'+n]/d.Volumen_um3).mean(),Centroide_medio=d['Centroide_X' if axx=='ML' else 'Centroide_Y'].mean()))
    pd.DataFrame(profiles).to_csv(OUT/'perfil_2D_grupo.csv',index=False)
    # Quadrants descriptives.
    quad=[]
    for (g,r,s),d in sec.groupby(['Grupo','Region','Sector']):
        dens=d.Densidad_celulas_mm3.astype(float);fracv=d.Volumen_um3.astype(float)/d.groupby(lambda x: True).Volumen_um3.transform('sum') if False else None
        # Merge per-animal total volume.
        t=base[(base.Grupo==g)&(base.Region==r)][['Animal_ID','Volumen_um3']]
        z=d.merge(t,on='Animal_ID',suffixes=('','_total'))
        quad.append(dict(Grupo=g,Region=r,Sector=s,n=len(d),Densidad_media=dens.mean(),Densidad_mediana=dens.median(),Fraccion_volumen_mediana=(z.Volumen_um3/z.Volumen_um3_total).median(),Fraccion_objetos_mediana=(z.Celulas_cFos/z.merge(base[(base.Grupo==g)&(base.Region==r)][['Animal_ID','Celulas_cFos']],on='Animal_ID',suffixes=('','_total')).Celulas_cFos_total).median()))
    pd.DataFrame(quad).to_csv(OUT/'cuadrantes_descriptivos.csv',index=False)
    # Bins 2D shape analysis: equal animal probability maps, Hellinger, residualized PERMANOVA.
    # Bins contain Celulas and Bin_X/Bin_Y. Sum stacks per animal.
    bins['Celulas']=pd.to_numeric(bins['Celulas'],errors='coerce').fillna(0)
    shape_rows=[];shape_sens=[]
    from scipy.ndimage import gaussian_filter
    for reg in ['BLAa','BLAp']:
        bb=bins[bins.Region==reg].groupby(['Animal_ID','Experimento','Animal','Grupo','Bin_X','Bin_Y'],as_index=False).Celulas.sum()
        meta=base[base.Region==reg][['Animal_ID','Experimento','Grupo','Bregma_medio']].drop_duplicates().sort_values('Animal_ID').reset_index(drop=True)
        for sig in [.75,1.5,2.5,4.0]:
            feats=[]
            for _,m in meta.iterrows():
                z=bb[bb.Animal_ID==m.Animal_ID];mat=np.zeros((30,30),float)
                for _,rr in z.iterrows():mat[int(rr.Bin_Y),int(rr.Bin_X)]+=rr.Celulas
                mat=gaussian_filter(mat,sig,mode='reflect');mat=mat/mat.sum() if mat.sum() else mat
                feats.append(np.sqrt(mat.ravel()))
            F=np.vstack(feats)
            for tr in ['SCOP-HB','LiCl']:
                d=meta[meta.Grupo.isin([tr,'Vehículo'])].copy();keep=[e for e,z in d.groupby('Experimento') if set(z.Grupo)=={tr,'Vehículo'}];d=d[d.Experimento.isin(keep)].copy().sort_values('Animal_ID').reset_index(drop=True)
                inds=[meta.index[meta.Animal_ID==aid][0] for aid in d.Animal_ID];Y=F[inds]
                # Residualize on intercept + experiment dummies + AP, then test treatment added via multivariate RSS.
                T=(d.Grupo==tr).astype(float).to_numpy();ex=pd.get_dummies(d.Experimento,drop_first=True,dtype=float);Z=np.column_stack([np.ones(len(d))]+([ex.to_numpy()] if ex.shape[1] else [])+[((d.Bregma_medio-d.Bregma_medio.mean())/(d.Bregma_medio.std(ddof=0) or 1)).to_numpy()])
                if ex.shape[1]:Z=np.column_stack([np.ones(len(d)),ex.to_numpy(),((d.Bregma_medio-d.Bregma_medio.mean())/(d.Bregma_medio.std(ddof=0) or 1)).to_numpy()])
                else:Z=np.column_stack([np.ones(len(d)),((d.Bregma_medio-d.Bregma_medio.mean())/(d.Bregma_medio.std(ddof=0) or 1)).to_numpy()])
                Pz=Z@np.linalg.pinv(Z);RY=Y-Pz@Y;rss0=float(np.sum(RY*RY));rankz=np.linalg.matrix_rank(Z)
                def statfun(t):
                    rt=t-Pz@t;den=float(rt@rt)
                    if den<=1e-12:return 0.0,0.0
                    coef=(rt@RY)/den;proj=rt[:,None]*coef[None,:];sse=float(np.sum(proj*proj));rss1=max(rss0-sse,1e-15);df2=max(len(d)-(rankz+1),1);Fv=sse/(rss1/df2);return Fv,(sse/rss0 if rss0 else 0.0)
                fo,r2=statfun(T);per=[]
                for tp in strat_perms(d,T):per.append(statfun(tp)[0])
                pp=float(np.mean(np.array(per)>=fo-1e-12));shape_sens.append(dict(Sigma=sig,Region=reg,Tratamiento=tr,Pseudo_F=fo,R2_parcial=r2,p_exacta=pp,Permutaciones=len(per)))
    shs=pd.DataFrame(shape_sens);shs['p_Holm_16']=multipletests(shs.p_exacta,method='holm')[1];shs.to_csv(OUT/'forma_2D_sensibilidad.csv',index=False)
    main=shs[np.isclose(shs.Sigma,1.5)].copy();main['p_Holm']=multipletests(main.p_exacta,method='holm')[1];main.to_csv(OUT/'forma_2D.csv',index=False)
    # AP coverage per group/bregma.
    cov=q.groupby(['Region','Grupo','Bregma_mm']).Animal_ID.nunique().rename('N_animales').reset_index();cov.to_csv(OUT/'cobertura_AP.csv',index=False)
    # 3D maxima support from NPZ.
    npzs=list(ROOT.rglob('3D_*.npz'));maxrows=[]
    for f in npzs:
        name=f.stem
        if name.startswith('diferencia'):continue
        arr=np.load(f)
        if not {'densidad_celulas_mm3','conteos','volumen_um3'}.issubset(arr.files):continue
        parts=name.split('_');reg=parts[-1];group='_'.join(parts[1:-1]).replace('Veh_culo','Vehículo')
        D=arr['densidad_celulas_mm3'];N=arr['conteos'];V=arr['volumen_um3'];idx=np.unravel_index(np.nanargmax(D),D.shape);peak=float(D[idx]);ncell=int(N[idx]);v=float(V[idx]);
        xe,ye,ae=arr['x_edges'],arr['y_edges'],arr['ap_edges_mm'];xl,xh=xe[idx[0]],xe[idx[0]+1];yl,yh=ye[idx[1]],ye[idx[1]+1];al,ah=ae[idx[2]],ae[idx[2]+1]
        c=cells[(cells.Grupo==group)&(cells.Region==reg)];inside=c[(c[XC]>=xl)&(c[XC]<xh)&(c[YC]>=yl)&(c[YC]<yh)&(c.AP_mm>=al)&(c.AP_mm<ah)]
        # infer animals with volume at AP bin and any group stacks overlapping bin
        qq=q[(q.Grupo==group)&(q.Region==reg)].copy();qq['low']=qq.Bregma_mm-pd.to_numeric(qq.Espesor_derivado_um,errors='coerce')/2000;qq['high']=qq.Bregma_mm+pd.to_numeric(qq.Espesor_derivado_um,errors='coerce')/2000;volanimals=qq[(qq.high>al)&(qq.low<ah)].Animal_ID.nunique()
        maxrows.append(dict(Region=reg,Grupo=group,Pico_mm3=peak,Celulas_en_voxel=ncell,Volumen_um3=v,Animales_con_celulas=inside.Animal_ID.nunique(),Animales_con_volumen=int(volanimals),X_centro=(xl+xh)/2,Y_centro=(yl+yh)/2,AP_centro=(al+ah)/2))
    pd.DataFrame(maxrows).to_csv(OUT/'maximos_3D.csv',index=False)
    # Spearman centroid AP vs sampled bregma.
    assoc=[]
    for reg,d in a.groupby('Region'):
        rho,pv=stats.spearmanr(d.Bregma_medio,d.Centroide_AP_mm);assoc.append(dict(Region=reg,rho=rho,p=pv,n=len(d)))
    pd.DataFrame(assoc).to_csv(OUT/'asociacion_AP.csv',index=False)
    # Store primary data.
    a.to_csv(OUT/'datos_animal.csv',index=False);q.to_csv(OUT/'qc_stacks.csv',index=False);base.to_csv(OUT/'polaridad_animal.csv',index=False)
    # Manifest
    manifest={'input':str(ROOT),'run':S['Version'].iloc[0].to_dict(),'summary':dict(summary),'sha256':{}}
    for f in [ROOT/'resultados_cfos_2D_3D.xlsx',ROOT/'conteo_objetos.xlsx',ROOT/'LEEME_resultados.txt']:
        manifest['sha256'][f.name]=hashlib.sha256(f.read_bytes()).hexdigest()
    (OUT/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2,default=str))
    print(json.dumps(manifest,ensure_ascii=False,indent=2,default=str))
    print('\nGLOBAL\n',glob.to_string(index=False));print('\nDIR\n',dir2.to_string(index=False));print('\nCENT\n',cent.to_string(index=False));print('\nSHAPE\n',main.to_string(index=False));print('\n3D\n',pd.DataFrame(maxrows).to_string(index=False))


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', required=True, type=Path, help='Carpeta de la corrida exportada por Heatmap')
    parser.add_argument('--output', required=True, type=Path, help='Carpeta de resultados estadísticos')
    args=parser.parse_args()
    run(args.input, args.output)


if __name__ == "__main__":
    main()
